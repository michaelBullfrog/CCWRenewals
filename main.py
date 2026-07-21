import os
import requests
import pandas as pd
import json
from pathlib import Path
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo


# =========================
# CONFIGURATION
# =========================

TOKEN_URL = "https://id.cisco.com/oauth2/default/v1/token"
GRAPHQL_URL = "https://capi.cisco.com/commerce/apis"

# Cisco credentials from environment variables
MARKETS = [
    {
        "market": "US",
        "client_id": os.getenv("CISCO_US_CLIENT_ID"),
        "client_secret": os.getenv("CISCO_US_CLIENT_SECRET"),
    },
    {
        "market": "Canada",
        "client_id": os.getenv("CISCO_CANADA_CLIENT_ID"),
        "client_secret": os.getenv("CISCO_CANADA_CLIENT_SECRET"),
    },
]

# Microsoft Graph / OneDrive environment variables
MS_TENANT_ID = os.getenv("MS_TENANT_ID")
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
ONEDRIVE_USER_ID = os.getenv("ONEDRIVE_USER_ID")
ONEDRIVE_FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH")
# SharePoint List tracker variables
# SHAREPOINT_SITE_ID can be the Graph site ID like:
# invoip-my.sharepoint.com,siteCollectionId,siteId
SHAREPOINT_SITE_ID = os.getenv("SHAREPOINT_SITE_ID")
SHAREPOINT_LIST_ID = os.getenv("SHAREPOINT_LIST_ID")

# These are SharePoint internal field names. The script will try to infer them from existing rows.
# If inference fails, set these in Render.
SP_SUBSCRIPTION_ID_FIELD = os.getenv("SP_SUBSCRIPTION_ID_FIELD", "")
SP_END_CUSTOMER_NAME_FIELD = os.getenv("SP_END_CUSTOMER_NAME_FIELD", "")

# Power BI semantic model refresh variables
POWERBI_TENANT_ID = os.getenv("POWERBI_TENANT_ID", MS_TENANT_ID)
POWERBI_CLIENT_ID = os.getenv("POWERBI_CLIENT_ID", MS_CLIENT_ID)
POWERBI_CLIENT_SECRET = os.getenv("POWERBI_CLIENT_SECRET", MS_CLIENT_SECRET)
POWERBI_GROUP_ID = os.getenv("POWERBI_GROUP_ID", "me")
POWERBI_DATASET_ID = os.getenv("POWERBI_DATASET_ID")

# Ribbit PostgreSQL ingestion endpoint
RIBBIT_BACKEND_URL = os.getenv("RIBBIT_BACKEND_URL", "").rstrip("/")
CCWR_INGEST_API_KEY = os.getenv("CCWR_INGEST_API_KEY", "")
RIBBIT_UPLOAD_TIMEOUT_SECONDS = int(os.getenv("RIBBIT_UPLOAD_TIMEOUT_SECONDS", "180"))

# Optional overrides
MASTER_START_DATE = os.getenv("MASTER_START_DATE", "2020-01-01")
WINDOW_DAYS = int(os.getenv("WINDOW_DAYS", "15"))
PAGE_SIZE = int(os.getenv("PAGE_SIZE", "100"))
MAX_PAGES_PER_WINDOW = int(os.getenv("MAX_PAGES_PER_WINDOW", "1"))

OUTPUT_FOLDER = Path("cisco_exports")
MASTER_CSV = OUTPUT_FOLDER / "cisco_master_subscription_renewals_all_statuses.csv"
RAW_JSON = OUTPUT_FOLDER / "raw_master_subscription_response_all_statuses.json"


# =========================
# VALIDATION
# =========================

def validate_environment():
    required = {
        "CISCO_US_CLIENT_ID": os.getenv("CISCO_US_CLIENT_ID"),
        "CISCO_US_CLIENT_SECRET": os.getenv("CISCO_US_CLIENT_SECRET"),
        "CISCO_CANADA_CLIENT_ID": os.getenv("CISCO_CANADA_CLIENT_ID"),
        "CISCO_CANADA_CLIENT_SECRET": os.getenv("CISCO_CANADA_CLIENT_SECRET"),
        "MS_TENANT_ID": MS_TENANT_ID,
        "MS_CLIENT_ID": MS_CLIENT_ID,
        "MS_CLIENT_SECRET": MS_CLIENT_SECRET,
        "ONEDRIVE_USER_ID": ONEDRIVE_USER_ID,
        "ONEDRIVE_FILE_PATH": ONEDRIVE_FILE_PATH,
        "SHAREPOINT_SITE_ID": SHAREPOINT_SITE_ID,
        "SHAREPOINT_LIST_ID": SHAREPOINT_LIST_ID,
        "RIBBIT_BACKEND_URL": RIBBIT_BACKEND_URL,
        "CCWR_INGEST_API_KEY": CCWR_INGEST_API_KEY,
    }

    missing = [key for key, value in required.items() if not value]
    if missing:
        raise RuntimeError(f"Missing required environment variables: {missing}")


# =========================
# CISCO AUTH
# =========================

def get_cisco_token(client_id, client_secret):
    response = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=60,
    )

    print("Cisco token status:", response.status_code)
    if response.status_code != 200:
        print(response.text)

    response.raise_for_status()
    return response.json()["access_token"]


# =========================
# GRAPHQL HELPER
# =========================

def run_graphql(token, query, client_id):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
        "client_id": client_id,
        "Client-Id": client_id,
        "X-Client-Id": client_id,
    }

    response = requests.post(
        GRAPHQL_URL,
        headers=headers,
        json={"query": query},
        timeout=120,
    )

    print("GraphQL status:", response.status_code)
    if response.status_code != 200:
        print(response.text)

    response.raise_for_status()
    data = response.json()

    if "errors" in data:
        print("GraphQL returned errors:")
        print(json.dumps(data["errors"], indent=2))

    return data


# =========================
# DATE WINDOWS
# =========================

def build_date_windows(start_date_string):
    start = datetime.strptime(start_date_string, "%Y-%m-%d").date()
    today = date.today()

    windows = []
    current_start = start

    while current_start <= today:
        current_end = current_start + timedelta(days=WINDOW_DAYS - 1)
        if current_end > today:
            current_end = today

        windows.append((current_start.strftime("%Y-%m-%d"), current_end.strftime("%Y-%m-%d")))
        current_start = current_end + timedelta(days=1)

    return windows


# =========================
# SEARCH SUBSCRIPTIONS
# =========================

def build_search_subscription_query(from_date, to_date, page):
    return f"""
    query SearchSubscription {{
        searchSubscription(
            input: {{
                mySubscriptionSearchCriteria: [
                    {{
                        mySubscriptionSearchKey: FROM_DATE
                        mySubscriptionSearchValue: "{from_date}"
                    }}
                    {{
                        mySubscriptionSearchKey: TO_DATE
                        mySubscriptionSearchValue: "{to_date}"
                    }}
                ]
                pagination: {{ page: {page}, pageSize: {PAGE_SIZE}, sortOrder: ASC }}
            }}
        ) {{
            businessStatus
            messages {{
                code
                description
                severity
                expecting
                exceptionMsg
            }}
            objects {{
                id
                parties {{
                    id
                    type
                    channelType
                    partnerType
                    name
                }}
                mySubscriptionCharacteristics {{
                    hasAutoRenewal
                    startDate
                    endDate
                    nextTrueForwardDate
                    renewalDate
                    mySubscriptionProvisioningStatus
                    billingModel
                    billingPreference
                    hasOverConsumption
                    mySubscriptionStatus
                    accountType
                    isAutoRenewalRequired
                    entitlementType
                    activationDate
                    initialTerm {{
                        measurement
                        unitOfMeasure
                    }}
                }}
            }}
        }}
    }}
    """


def search_subscriptions_for_window(token, client_id, market, from_date, to_date):
    all_objects = []
    raw_pages = []

    for page in range(1, MAX_PAGES_PER_WINDOW + 1):
        print(f"\n[{market}] Searching window {from_date} to {to_date} - page {page}")

        query = build_search_subscription_query(from_date=from_date, to_date=to_date, page=page)
        data = run_graphql(token=token, query=query, client_id=client_id)

        raw_pages.append({
            "market": market,
            "from_date": from_date,
            "to_date": to_date,
            "page": page,
            "response": data,
        })

        search_result = data.get("data", {}).get("searchSubscription", {})
        business_status = search_result.get("businessStatus")
        messages = search_result.get("messages", [])
        objects = search_result.get("objects", []) or []

        print("Business status:", business_status)

        if messages:
            print("Messages:")
            print(json.dumps(messages, indent=2))

        print(f"Objects returned: {len(objects)}")

        if len(objects) >= PAGE_SIZE:
            print(
                f"WARNING: [{market}] Window {from_date} to {to_date} returned "
                f"{PAGE_SIZE}. This may be capped. Consider reducing WINDOW_DAYS."
            )

        all_objects.extend(objects)

        if len(objects) < PAGE_SIZE:
            break

    return all_objects, raw_pages


# =========================
# FLATTEN RESULTS
# =========================

def get_party(parties, party_type):
    for party in parties:
        if party.get("type") == party_type:
            return {
                "id": party.get("id", ""),
                "name": party.get("name", ""),
                "channelType": party.get("channelType", ""),
                "partnerType": party.get("partnerType", ""),
            }

    return {"id": "", "name": "", "channelType": "", "partnerType": ""}


def flatten_subscription_search_results(objects, market):
    rows = []

    for sub in objects:
        chars = sub.get("mySubscriptionCharacteristics", {}) or {}
        parties = sub.get("parties", []) or []

        bill_to = get_party(parties, "BILL_TO")
        reseller = get_party(parties, "RESELLER")
        end_customer = get_party(parties, "END_CUSTOMER")
        ship_to = get_party(parties, "SHIP_TO")
        initial_term = chars.get("initialTerm", {}) or {}

        rows.append({
            "Market": market,
            "Subscription ID": sub.get("id", ""),
            "Bill To ID": bill_to["id"],
            "Bill To Name": bill_to["name"],
            "Reseller ID": reseller["id"],
            "Reseller Name": reseller["name"],
            "End Customer ID": end_customer["id"],
            "End Customer Name": end_customer["name"],
            "Ship To ID": ship_to["id"],
            "Ship To Name": ship_to["name"],
            "Start Date": chars.get("startDate", ""),
            "End Date": chars.get("endDate", ""),
            "Renewal Date": chars.get("renewalDate", ""),
            "Next True Forward Date": chars.get("nextTrueForwardDate", ""),
            "Subscription Status": chars.get("mySubscriptionStatus", ""),
            "Provisioning Status": chars.get("mySubscriptionProvisioningStatus", ""),
            "Billing Model": chars.get("billingModel", ""),
            "Billing Preference": chars.get("billingPreference", ""),
            "Has Auto Renewal": chars.get("hasAutoRenewal", ""),
            "Auto Renewal Required": chars.get("isAutoRenewalRequired", ""),
            "Has Over Consumption": chars.get("hasOverConsumption", ""),
            "Entitlement Type": chars.get("entitlementType", ""),
            "Account Type": chars.get("accountType", ""),
            "Activation Date": chars.get("activationDate", ""),
            "Initial Term Measurement": initial_term.get("measurement", ""),
            "Initial Term UOM": initial_term.get("unitOfMeasure", ""),
        })

    return rows


# =========================
# DASHBOARD FIELDS
# =========================

def add_dashboard_fields(df):
    today = pd.Timestamp.today().normalize()

    df["Renewal Date Parsed"] = pd.to_datetime(df["Renewal Date"], errors="coerce")
    df["End Date Parsed"] = pd.to_datetime(df["End Date"], errors="coerce")
    df["Start Date Parsed"] = pd.to_datetime(df["Start Date"], errors="coerce")

    df["Dashboard Renewal Date"] = df["Renewal Date Parsed"].fillna(df["End Date Parsed"])
    df["Days Until Renewal"] = (df["Dashboard Renewal Date"] - today).dt.days

    def renewal_window(days):
        if pd.isna(days):
            return "No Date"
        if days < 0:
            return "Past Due"
        if days <= 30:
            return "0-30 Days"
        if days <= 60:
            return "31-60 Days"
        if days <= 90:
            return "61-90 Days"
        return "91+ Days"

    df["Renewal Window"] = df["Days Until Renewal"].apply(renewal_window)

    def bucket(days):
        if pd.isna(days):
            return "No Date"
        if days < 0:
            return "Expired"
        if days <= 30:
            return "0-30 Days"
        if days <= 60:
            return "31-60 Days"
        if days <= 90:
            return "61-90 Days"
        if days <= 180:
            return "91-180 Days"
        if days <= 365:
            return "181-365 Days"
        return "365+ Days"

    df["Renewal Bucket"] = df["Days Until Renewal"].apply(bucket)

    def risk_level(days):
        if pd.isna(days):
            return "Unknown"
        if days < 0:
            return "Expired"
        if days <= 30:
            return "High"
        if days <= 90:
            return "Medium"
        return "Low"

    df["Renewal Risk"] = df["Days Until Renewal"].apply(risk_level)

    eastern_now = datetime.now(ZoneInfo("America/New_York"))
    utc_now = datetime.now(ZoneInfo("UTC"))

    df["Last Refreshed"] = eastern_now.strftime("%m/%d/%Y %I:%M:%S %p ET")
    df["Last Refreshed UTC"] = utc_now.strftime("%m/%d/%Y %I:%M:%S %p UTC")

    return df


# =========================
# MICROSOFT GRAPH HELPERS
# =========================

def get_graph_token():
    token_url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"

    response = requests.post(
        token_url,
        data={
            "client_id": MS_CLIENT_ID,
            "client_secret": MS_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=60,
    )

    print("Graph token status:", response.status_code)
    if response.status_code != 200:
        print(response.text)

    response.raise_for_status()
    return response.json()["access_token"]


def upload_csv_to_onedrive(token, local_file_path):
    upload_url = (
        f"https://graph.microsoft.com/v1.0/users/{ONEDRIVE_USER_ID}"
        f"/drive/root:/{ONEDRIVE_FILE_PATH}:/content"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "text/csv",
    }

    with open(local_file_path, "rb") as file_data:
        response = requests.put(upload_url, headers=headers, data=file_data, timeout=120)

    print("OneDrive CSV upload status:", response.status_code)
    if response.status_code not in (200, 201):
        print(response.text)

    response.raise_for_status()
    result = response.json()
    print("Uploaded CSV to OneDrive:", result.get("name"))
    print("CSV OneDrive URL:", result.get("webUrl"))



# =========================
# SHAREPOINT LIST TRACKER
# =========================

def graph_get(token, url):
    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=120,
    )
    print("Graph GET status:", response.status_code)
    if response.status_code != 200:
        print(response.text)
    response.raise_for_status()
    return response.json()


def graph_post(token, url, payload):
    response = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=120,
    )
    if response.status_code not in (200, 201, 202):
        print("Graph POST status:", response.status_code)
        print(response.text)
    response.raise_for_status()
    return response.json() if response.text else {}


def graph_patch(token, url, payload):
    response = requests.patch(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=120,
    )
    if response.status_code not in (200, 204):
        print("Graph PATCH status:", response.status_code)
        print(response.text)
    response.raise_for_status()
    return response.json() if response.text else {}


def get_all_sharepoint_tracker_items(token):
    """Return all SharePoint list items with their fields."""
    items = []
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}"
        f"/lists/{SHAREPOINT_LIST_ID}/items?expand=fields&top=999"
    )

    while url:
        data = graph_get(token, url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return items


def choose_field_name(items, env_value, candidates, fallback):
    """
    Choose a SharePoint internal field name.
    If env_value is set, use it. Otherwise infer from existing list item fields.
    """
    if env_value:
        return env_value

    observed_keys = set()
    for item in items:
        observed_keys.update((item.get("fields") or {}).keys())

    for candidate in candidates:
        if candidate in observed_keys:
            return candidate

    return fallback


def sync_cisco_subscriptions_to_sharepoint_list(df, token):
    """
    Syncs Subscription ID + End Customer Name from Cisco data to a SharePoint List.
    It creates missing rows and updates customer names.
    It does not overwrite sales fields such as Quoted, Signed, Sales Rep, Notes, etc.
    """
    print("\nSyncing Cisco subscriptions to SharePoint List tracker...")

    existing_items = get_all_sharepoint_tracker_items(token)
    print(f"SharePoint List existing tracked subscriptions: {len(existing_items)}")

    subscription_field = choose_field_name(
        existing_items,
        SP_SUBSCRIPTION_ID_FIELD,
        [
            "SubscriptionID",
            "Subscription_x0020_ID",
            "Subscription ID",
            "SubID",
            "Sub_x0020_ID",
            "Title",
        ],
        "Title",
    )

    customer_field = choose_field_name(
        existing_items,
        SP_END_CUSTOMER_NAME_FIELD,
        [
            "EndCustomerName",
            "End_x0020_Customer_x0020_Name",
            "End Customer Name",
            "CustomerName",
            "Customer_x0020_Name",
        ],
        "EndCustomerName",
    )

    print(f"Using SharePoint Subscription ID field: {subscription_field}")
    print(f"Using SharePoint End Customer Name field: {customer_field}")

    existing_by_sub_id = {}
    for item in existing_items:
        fields = item.get("fields") or {}
        sub_id = str(fields.get(subscription_field, "")).strip()
        if sub_id:
            existing_by_sub_id[sub_id] = item

    cisco_lookup = (
        df[["Subscription ID", "End Customer Name"]]
        .dropna(subset=["Subscription ID"])
        .drop_duplicates(subset=["Subscription ID"])
        .copy()
    )

    cisco_lookup["Subscription ID"] = (
        cisco_lookup["Subscription ID"].fillna("").astype(str).str.strip()
    )
    cisco_lookup["End Customer Name"] = (
        cisco_lookup["End Customer Name"].fillna("").astype(str).str.strip()
    )

    create_url = (
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}"
        f"/lists/{SHAREPOINT_LIST_ID}/items"
    )

    created = 0
    updated = 0
    skipped = 0

    for _, row in cisco_lookup.iterrows():
        sub_id = row["Subscription ID"]
        customer_name = row["End Customer Name"]

        if not sub_id:
            skipped += 1
            continue

        if sub_id not in existing_by_sub_id:
            fields_payload = {
                subscription_field: sub_id,
                customer_field: customer_name,
            }

            # If the subscription field is not Title, still set Title for easier SharePoint browsing.
            if subscription_field != "Title":
                fields_payload["Title"] = sub_id

            payload = {"fields": fields_payload}
            try:
                graph_post(token, create_url, payload)
                created += 1
                if created % 25 == 0:
                    print(f"SharePoint tracker rows created so far: {created}")
            except Exception as e:
                print(f"Failed creating SharePoint tracker row for {sub_id}: {e}")
                skipped += 1
            continue

        item = existing_by_sub_id[sub_id]
        fields = item.get("fields") or {}
        current_customer = str(fields.get(customer_field, "")).strip()

        if customer_name and current_customer != customer_name:
            patch_url = (
                f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}"
                f"/lists/{SHAREPOINT_LIST_ID}/items/{item['id']}/fields"
            )
            try:
                graph_patch(token, patch_url, {customer_field: customer_name})
                updated += 1
            except Exception as e:
                print(f"Failed updating customer name for {sub_id}: {e}")
                skipped += 1

    print(f"SharePoint tracker rows created: {created}")
    print(f"SharePoint tracker customer names updated: {updated}")
    print(f"SharePoint tracker rows skipped: {skipped}")
    print("SharePoint List tracker sync complete.")


# =========================
# POWER BI SEMANTIC MODEL REFRESH
# =========================

def get_powerbi_token():
    token_url = f"https://login.microsoftonline.com/{POWERBI_TENANT_ID}/oauth2/v2.0/token"

    response = requests.post(
        token_url,
        data={
            "client_id": POWERBI_CLIENT_ID,
            "client_secret": POWERBI_CLIENT_SECRET,
            "scope": "https://analysis.windows.net/powerbi/api/.default",
            "grant_type": "client_credentials",
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=60,
    )

    print("Power BI token status:", response.status_code)
    if response.status_code != 200:
        print(response.text)

    response.raise_for_status()
    return response.json()["access_token"]


def trigger_powerbi_semantic_model_refresh():
    if not POWERBI_DATASET_ID:
        print("POWERBI_DATASET_ID not set. Skipping Power BI semantic model refresh.")
        return

    print("\nTriggering Power BI semantic model refresh...")
    token = get_powerbi_token()

    if not POWERBI_GROUP_ID or POWERBI_GROUP_ID.lower() == "me":
        refresh_url = (
            f"https://api.powerbi.com/v1.0/myorg/datasets/"
            f"{POWERBI_DATASET_ID}/refreshes"
        )
    else:
        refresh_url = (
            f"https://api.powerbi.com/v1.0/myorg/groups/"
            f"{POWERBI_GROUP_ID}/datasets/{POWERBI_DATASET_ID}/refreshes"
        )

    response = requests.post(
        refresh_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"notifyOption": "NoNotification"},
        timeout=60,
    )

    print("Power BI refresh trigger status:", response.status_code)
    if response.status_code not in (200, 202):
        print(response.text)

    response.raise_for_status()
    print("Power BI semantic model refresh triggered successfully.")


# =========================
# RIBBIT DATABASE UPLOAD
# =========================

def _clean_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def upload_snapshot_to_ribbit(df):
    refreshed_at = datetime.now(ZoneInfo("UTC")).isoformat()
    sync_id = f"ccwr-{datetime.now(ZoneInfo('UTC')).strftime('%Y%m%dT%H%M%SZ')}"

    records = []
    for _, row in df.iterrows():
        records.append({
            "subscription_id": _clean_value(row.get("Subscription ID")),
            "market": _clean_value(row.get("Market")),
            "subscription_status": _clean_value(row.get("Subscription Status")),
            "renewal_date": _clean_value(row.get("Renewal Date Parsed")),
            "dashboard_renewal_date": _clean_value(row.get("Dashboard Renewal Date")),
            "renewal_bucket": _clean_value(row.get("Renewal Bucket")),
            "renewal_window": _clean_value(row.get("Renewal Window")),
            "renewal_risk": _clean_value(row.get("Renewal Risk")),
            "days_until_renewal": _clean_value(row.get("Days Until Renewal")),
            "end_customer_name": _clean_value(row.get("End Customer Name")),
            "end_customer_id": _clean_value(row.get("End Customer ID")),
            "reseller_name": _clean_value(row.get("Reseller Name")),
            "reseller_id": _clean_value(row.get("Reseller ID")),
            "bill_to_name": _clean_value(row.get("Bill To Name")),
            "bill_to_id": _clean_value(row.get("Bill To ID")),
            "has_auto_renewal": _clean_value(row.get("Has Auto Renewal")),
            "provisioning_status": _clean_value(row.get("Provisioning Status")),
            "billing_model": _clean_value(row.get("Billing Model")),
        })

    payload = {
        "sync_id": sync_id,
        "refreshed_at": refreshed_at,
        "replace_snapshot": True,
        "records": records,
    }
    url = f"{RIBBIT_BACKEND_URL}/api/ingest/ccwr-renewals"
    print(f"\nUploading {len(records)} renewals to Ribbit: {url}")
    response = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {CCWR_INGEST_API_KEY}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=RIBBIT_UPLOAD_TIMEOUT_SECONDS,
    )
    print("Ribbit upload status:", response.status_code)
    if response.status_code >= 400:
        print(response.text)
    response.raise_for_status()
    result = response.json()
    print("Ribbit renewal sync completed:")
    print(json.dumps(result, indent=2))
    return result


# COVERAGE SUMMARY
# =========================

def print_coverage_summary(df):
    print("\n" + "=" * 80)
    print("COVERAGE SUMMARY")
    print("=" * 80)

    print("Unique subscriptions:", len(df))

    print("\nCount by Market:")
    print(df["Market"].value_counts(dropna=False))

    print("\nCount by Subscription Status:")
    print(df["Subscription Status"].value_counts(dropna=False))

    print("\nCount by Market and Subscription Status:")
    print(pd.crosstab(df["Market"], df["Subscription Status"]))

    print("\nEarliest Start Date:", df["Start Date Parsed"].min())
    print("Latest Start Date:", df["Start Date Parsed"].max())

    print("\nEarliest Renewal Date:", df["Dashboard Renewal Date"].min())
    print("Latest Renewal Date:", df["Dashboard Renewal Date"].max())

    print("\nCount by Bill To Name:")
    print(df["Bill To Name"].value_counts(dropna=False))

    print("\nTop 20 Resellers:")
    print(df["Reseller Name"].value_counts(dropna=False).head(20))

    print("\nRenewal Window Count:")
    print(df["Renewal Window"].value_counts(dropna=False))

    print("\nRenewal Bucket Count:")
    print(df["Renewal Bucket"].value_counts(dropna=False))

    print("\nRenewal Risk Count:")
    print(df["Renewal Risk"].value_counts(dropna=False))


# =========================
# MAIN
# =========================

def main():
    validate_environment()
    OUTPUT_FOLDER.mkdir(exist_ok=True)

    windows = build_date_windows(MASTER_START_DATE)

    print(f"\nTotal date windows per market: {len(windows)}")
    print(f"Master start date: {MASTER_START_DATE}")
    print(f"Today: {date.today().strftime('%Y-%m-%d')}")
    print(f"Window days: {WINDOW_DAYS}")
    print(f"Page size: {PAGE_SIZE}")
    print("Status filter: NONE")
    print("Party / Bill-To / Disti filter: NONE")
    print(f"CSV OneDrive path: {ONEDRIVE_FILE_PATH}")
    print(f"SharePoint Site ID: {SHAREPOINT_SITE_ID}")
    print(f"SharePoint List ID: {SHAREPOINT_LIST_ID}")
    print(f"Power BI Group ID: {POWERBI_GROUP_ID}")
    print(f"Power BI Dataset ID: {POWERBI_DATASET_ID or 'not configured'}")
    print(f"Ribbit Backend URL: {RIBBIT_BACKEND_URL}")

    all_rows = []
    all_raw = []

    for market_config in MARKETS:
        market = market_config["market"]
        client_id = market_config["client_id"]
        client_secret = market_config["client_secret"]

        print("\n" + "#" * 80)
        print(f"STARTING MARKET: {market}")
        print("No Bill-To / Disti filter will be used.")
        print("#" * 80)

        try:
            print(f"\n[{market}] Getting Cisco token...")
            token = get_cisco_token(client_id, client_secret)
        except Exception as e:
            print(f"[{market}] Failed to get Cisco token: {e}")
            continue

        market_objects = []

        for index, (from_date, to_date) in enumerate(windows, start=1):
            print("\n" + "=" * 80)
            print(f"[{market}] Window {index}/{len(windows)}: {from_date} to {to_date}")
            print("=" * 80)

            try:
                objects, raw_pages = search_subscriptions_for_window(
                    token=token,
                    client_id=client_id,
                    market=market,
                    from_date=from_date,
                    to_date=to_date,
                )
                market_objects.extend(objects)
                all_raw.extend(raw_pages)
            except requests.exceptions.HTTPError as e:
                print(f"[{market}] HTTP error for window {from_date} to {to_date}: {e}")
                continue
            except Exception as e:
                print(f"[{market}] Unexpected error for window {from_date} to {to_date}: {e}")
                continue

        print(f"\n[{market}] Raw rows returned before flattening: {len(market_objects)}")
        all_rows.extend(flatten_subscription_search_results(objects=market_objects, market=market))

    with open(RAW_JSON, "w", encoding="utf-8") as f:
        json.dump(all_raw, f, indent=2)

    print("\n" + "=" * 80)
    print("BUILDING COMBINED CSV")
    print("=" * 80)

    df = pd.DataFrame(all_rows)

    if df.empty:
        print("No subscription results returned.")
        print(f"Raw response saved to: {RAW_JSON}")
        return

    before_dedupe = len(df)
    df = df.drop_duplicates(subset=["Market", "Subscription ID"])
    after_dedupe = len(df)

    print(f"Rows before dedupe: {before_dedupe}")
    print(f"Rows after dedupe: {after_dedupe}")

    if "Subscription Status" in df.columns:
        df["Subscription Status"] = df["Subscription Status"].astype(str).str.upper().str.strip()

    df = add_dashboard_fields(df)

    df = df.sort_values(
        by=["Market", "Dashboard Renewal Date", "End Customer Name"],
        ascending=[True, True, True],
    )

    df.to_csv(MASTER_CSV, index=False)

    print(f"\nSaved combined master renewal file locally to: {MASTER_CSV}")
    print(f"Saved raw API response locally to: {RAW_JSON}")

    print_coverage_summary(df)

    upload_snapshot_to_ribbit(df)

    print("\nGetting Microsoft Graph token...")
    graph_token = get_graph_token()

    print("\nUploading CSV to OneDrive...")
    upload_csv_to_onedrive(graph_token, MASTER_CSV)

    sync_cisco_subscriptions_to_sharepoint_list(df, graph_token)

    trigger_powerbi_semantic_model_refresh()

    print("\nPreview:")
    preview_columns = [
        "Market",
        "End Customer Name",
        "Subscription ID",
        "Bill To Name",
        "Reseller Name",
        "Start Date",
        "End Date",
        "Renewal Date",
        "Days Until Renewal",
        "Renewal Window",
        "Renewal Bucket",
        "Renewal Risk",
        "Subscription Status",
        "Provisioning Status",
        "Billing Model",
        "Last Refreshed",
        "Last Refreshed UTC",
    ]

    available_preview_columns = [col for col in preview_columns if col in df.columns]
    print(df[available_preview_columns].head(50))

    print("\nDone.")


if __name__ == "__main__":
    main()
