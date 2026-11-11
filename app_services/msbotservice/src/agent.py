"""teams-bot-final-product — all backend logic lives here, by design: bot wiring,
storage (Blob/SharePoint toggle), the request-tracking table, the queue hand-off,
the call to the existing generate_otc Function (via APIM), and proactive Teams
messaging. frontend.py is presentation-only and calls into this module's HTTP
routes (wired in function_app.py); nothing here depends on frontend.py.

EVERYTHING CONFIGURABLE lives in the CONFIG section below — one place, env vars
only, no hardcoded URLs/names/keys. See Constants.py for the env var names and
defaults.

THE CALL TO generate_otc: the auth pattern (Ocp-Apim-Subscription-Key header,
aiohttp.ClientSession) is ported from project-2-best-practice-solution's
CommonHandler.py — but that project calls a *different* APIM instance, so its
URL shape ("/otc/generate") turned out NOT to transfer directly here. The actual
path for *this* APIM instance was confirmed by
reading the real operation definition (az rest GET .../apis/otc-orchestrator/
operations) rather than assumed: API path "otc" + operation urlTemplate
"/otc/generate" compose to "/otc/otc/generate" externally — see the comment at
the call site. The request payload shape is now partially CONFIRMED (not
assumed) via a live Application Insights traceback from generate_otc's own
code after it raised KeyError('investmentMemoBlobPath') on the originally
assumed nested shape — see _build_generate_otc_payload()'s docstring. The
response shape remains unconfirmed; everything here is built to fail loudly
(HTTP status + body logged, tracking row marked Failed with the real error)
rather than silently.

SHAREPOINT TOGGLE: USE_SHAREPOINT=true switches document storage to a SharePoint
document library via Microsoft Graph; USE_SHAREPOINT=false (default) uses Blob
Storage, the proven path from every other project here. The SharePoint branch is
implemented but UNVERIFIED — provisioning an actual SharePoint site/list and
granting this app's identity Graph permissions (Sites.ReadWrite.All or a scoped
site permission) hasn't happened yet; see upload_document()'s sharepoint branch.
"""

import json
import logging
import os
import sys
import uuid
from datetime import datetime, timedelta, timezone
from os import environ
from random import randint
from typing import Any

import aiohttp
import requests
from azure.core.exceptions import ResourceExistsError, ResourceNotFoundError
from azure.data.tables import TableClient
from azure.identity import ClientSecretCredential, DefaultAzureCredential, ManagedIdentityCredential
from azure.storage.blob import BlobSasPermissions, BlobServiceClient, ContentSettings, generate_blob_sas
from azure.storage.queue import QueueClient

from azure_functions_adapter import AzureFunctionsAdapter
from microsoft_agents.hosting.core import (
    Authorization,
    AgentApplication,
    TurnState,
    TurnContext,
    MemoryStorage,
)
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.activity import load_configuration_from_env

from Constants import ConstMessages, Defaults, EnvVariables

logger = logging.getLogger("teams_bot_final_product.agent")

agents_sdk_config = load_configuration_from_env(environ)

# ---------------------------------------------------------------------------
# CONFIG — every external dependency, one place, env vars only.
# ---------------------------------------------------------------------------
APIM_BASE_URL = environ.get(EnvVariables.APIM_BASE_URL, "").rstrip("/")
APIM_SUBSCRIPTION_KEY = environ.get(EnvVariables.APIM_SUBSCRIPTION_KEY, "")

STORAGE_ACCOUNT_NAME = environ.get(EnvVariables.STORAGE_ACCOUNT_NAME, "")
STORAGE_UPLOADS_CONTAINER = environ.get(EnvVariables.STORAGE_UPLOADS_CONTAINER, Defaults.STORAGE_UPLOADS_CONTAINER)
STORAGE_GENERATED_CONTAINER = environ.get(EnvVariables.STORAGE_GENERATED_CONTAINER, Defaults.STORAGE_GENERATED_CONTAINER)
AZURE_STORAGE_CONNECTION_STRING = environ.get(EnvVariables.AZURE_STORAGE_CONNECTION_STRING, "")
CONVERSATION_STORE_TABLE_NAME = environ.get(EnvVariables.CONVERSATION_STORE_TABLE_NAME, Defaults.CONVERSATION_STORE_TABLE_NAME)
TRACKING_TABLE_NAME = environ.get(EnvVariables.TRACKING_TABLE_NAME, Defaults.TRACKING_TABLE_NAME)
STORAGE_QUEUE_NAME = environ.get(EnvVariables.STORAGE_QUEUE_NAME, Defaults.STORAGE_QUEUE_NAME)
APP_BASE_URL = environ.get(EnvVariables.APP_BASE_URL, "").rstrip("/")
UPLOAD_SAS_TTL_MINUTES = int(environ.get(EnvVariables.UPLOAD_SAS_TTL_MINUTES, str(Defaults.UPLOAD_SAS_TTL_MINUTES)))

# The boolean toggle. Read once at import time; "true" (any case) -> SharePoint,
# anything else -> Blob Storage.
USE_SHAREPOINT = environ.get(EnvVariables.USE_SHAREPOINT, "false").strip().lower() == "true"
SHAREPOINT_SITE_URL = environ.get(EnvVariables.SHAREPOINT_SITE_URL, "")
SHAREPOINT_LIST_NAME = environ.get(EnvVariables.SHAREPOINT_LIST_NAME, Defaults.SHAREPOINT_LIST_NAME)
SHAREPOINT_DRIVE_ID = environ.get(EnvVariables.SHAREPOINT_DRIVE_ID, "")
# Raw Graph site ID ("host,siteGuid,webGuid"), needed for the Lists API -
# distinct from SHAREPOINT_SITE_URL, which is just the human-readable URL.
SHAREPOINT_SITE_ID = environ.get(EnvVariables.SHAREPOINT_SITE_ID, "")

# The SharePoint site lives in a different Entra tenant than this Function App's
# UAMI (confirmed via tenant discovery) - a managed identity cannot cross tenant
# boundaries, so Graph calls for the SharePoint branch use this tenant-scoped app
# registration + client secret instead. Nothing else in this project uses a
# secret-based credential - this is the one deliberate exception, confined to
# _graph_token() below.
SHAREPOINT_TENANT_ID = environ.get(EnvVariables.SHAREPOINT_TENANT_ID, "")
SHAREPOINT_CLIENT_ID = environ.get(EnvVariables.SHAREPOINT_CLIENT_ID, "")
SHAREPOINT_CLIENT_SECRET = environ.get(EnvVariables.SHAREPOINT_CLIENT_SECRET, "")


def _sharepoint_configured() -> bool:
    """Whether the SharePoint *list* (tracking/visibility mirror) can be
    written to. Deliberately independent of USE_SHAREPOINT, which only
    selects where DOCUMENTS are stored (Blob vs SharePoint) - the tracking
    list is written in both storage modes whenever these four settings are
    present, so SharePoint stays a live record of every request regardless of
    which backend holds the actual files."""
    return bool(SHAREPOINT_SITE_ID and SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET)

# Same three settings the existing Container App (ca-otc-agent-dev) already uses
# for Bot Framework identity — reused as-is, not renamed.
MICROSOFT_APP_ID = environ.get("MicrosoftAppId", "") or environ.get("MICROSOFT_APP_ID", "")
MICROSOFT_APP_TENANT_ID = environ.get("MicrosoftAppTenantId", "") or environ.get("MICROSOFT_APP_TENANT_ID", "")
AZURE_CLIENT_ID = environ.get("AZURE_CLIENT_ID", "")

BOT_CONNECTOR_SCOPE = "https://api.botframework.com/.default"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"

if not APIM_BASE_URL:
    print("Warning: APIM_BASE_URL not set. generate_otc calls will fail.", file=sys.stderr)
if not STORAGE_ACCOUNT_NAME and not AZURE_STORAGE_CONNECTION_STRING:
    print("Warning: STORAGE_ACCOUNT_NAME not set. Storage calls will fail.", file=sys.stderr)

STORAGE = MemoryStorage()
CONNECTION_MANAGER = MsalConnectionManager(**agents_sdk_config)
ADAPTER = AzureFunctionsAdapter(connection_manager=CONNECTION_MANAGER)
AUTHORIZATION = Authorization(STORAGE, CONNECTION_MANAGER, **agents_sdk_config)


# ---------------------------------------------------------------------------
# Storage clients — same DefaultAzureCredential/UAMI pattern as every other
# project here (project-2-best-practice-solution/CommonHandler.py).
# ---------------------------------------------------------------------------
def _blob_service_client() -> BlobServiceClient:
    if AZURE_STORAGE_CONNECTION_STRING:
        return BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)
    return BlobServiceClient(f"https://{STORAGE_ACCOUNT_NAME}.blob.core.windows.net", credential=DefaultAzureCredential())


def _table_client(table_name: str) -> TableClient:
    if AZURE_STORAGE_CONNECTION_STRING:
        return TableClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING, table_name)
    return TableClient(f"https://{STORAGE_ACCOUNT_NAME}.table.core.windows.net", table_name, credential=DefaultAzureCredential())


def _queue_client() -> QueueClient:
    if AZURE_STORAGE_CONNECTION_STRING:
        return QueueClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING, STORAGE_QUEUE_NAME)
    return QueueClient(f"https://{STORAGE_ACCOUNT_NAME}.queue.core.windows.net", STORAGE_QUEUE_NAME, credential=DefaultAzureCredential())


# ---------------------------------------------------------------------------
# Request ID — "OTC-2026-004521": year + 6-digit reference number, assigned
# once at creation and FIXED for the request's lifetime (status is tracked
# separately via STATUS_* and shown via the badge/tracker, never by changing
# this ID). This is deliberate, not a simplification: the ID is also the
# storage path prefix for uploaded documents (Blob and SharePoint), the Table
# Storage RowKey, the search key, and the generated document's download link —
# if it changed mid-flight, every reference created under the old ID would
# break (orphaned uploads, a dead search result, a broken download/
# proactive-card link). The 6-digit suffix is random, not a true atomic
# sequence counter (would need its own concurrency-safe counter entity) — at
# CEO-demo volume the collision odds are negligible, and it still reads as a
# real enterprise reference number.
# ---------------------------------------------------------------------------
def build_request_id() -> str:
    year = datetime.now(timezone.utc).year
    suffix = randint(0, 999999)
    return f"OTC-{year}-{suffix:06d}"


# ---------------------------------------------------------------------------
# Conversation references (for proactive messaging) — identical pattern to
# project-2-best-practice-solution: saved per-user on every bot turn, re-saved
# under the request_id at submit time.
# ---------------------------------------------------------------------------
_USER_CONV_PARTITION = "user-conversation"


def save_user_conversation_reference(teams_user_id: str, conversation_reference: dict[str, Any]) -> None:
    client = _table_client(CONVERSATION_STORE_TABLE_NAME)
    try:
        client.create_table()
    except ResourceExistsError:
        pass
    client.upsert_entity(
        {
            "PartitionKey": _USER_CONV_PARTITION,
            "RowKey": teams_user_id,
            "conversationReferenceJson": json.dumps(conversation_reference, default=str),
            "updatedAt": datetime.now(timezone.utc).isoformat(),
        }
    )


def get_user_conversation_reference(teams_user_id: str) -> dict[str, Any] | None:
    client = _table_client(CONVERSATION_STORE_TABLE_NAME)
    try:
        entity = client.get_entity(_USER_CONV_PARTITION, teams_user_id)
    except ResourceNotFoundError:  # no row yet, or table not yet created - both 404
        return None
    return json.loads(entity["conversationReferenceJson"])


def save_request_conversation_reference(request_id: str, conversation_reference: dict[str, Any]) -> None:
    client = _table_client(CONVERSATION_STORE_TABLE_NAME)
    try:
        client.create_table()
    except ResourceExistsError:
        pass
    client.upsert_entity(
        {
            "PartitionKey": request_id,
            "RowKey": request_id,
            "conversationReferenceJson": json.dumps(conversation_reference, default=str),
            "createdAt": datetime.now(timezone.utc).isoformat(),
        }
    )


def get_request_conversation_reference(request_id: str) -> dict[str, Any] | None:
    client = _table_client(CONVERSATION_STORE_TABLE_NAME)
    try:
        entity = client.get_entity(request_id, request_id)
    except ResourceNotFoundError:  # no row yet, or table not yet created - both 404
        return None
    return json.loads(entity["conversationReferenceJson"])


# ---------------------------------------------------------------------------
# Tracking table — one row per request, updated as it moves through the
# pipeline. Backs both the live side-panel progress and the search-by-ID box
# in frontend.py.
# ---------------------------------------------------------------------------
def create_tracking_record(request_id: str, fields: dict[str, str], requested_by: str, teams_user_id: str) -> None:
    client = _table_client(TRACKING_TABLE_NAME)
    try:
        client.create_table()
    except ResourceExistsError:
        pass
    now = datetime.now(timezone.utc).isoformat()
    entity: dict[str, Any] = {
        "PartitionKey": "request",
        "RowKey": request_id,
        "requestId": request_id,
        "status": ConstMessages.STATUS_CREATED,
        "customerName": fields.get("customerName", ""),
        "requestType": fields.get("requestType", ""),
        "facilityType": fields.get("facilityType", ""),
        "region": fields.get("region", ""),
        "country": fields.get("country", ""),
        "requestedBy": requested_by,
        "teamsUserId": teams_user_id,
        "storageMode": "sharepoint" if USE_SHAREPOINT else "blob",
        "createdAt": now,
        "updatedAt": now,
        "downloadUrl": "",
        "fileName": "",
        "fileSizeKb": 0,
        "errorMessage": "",
        "sharepointListItemId": "",
    }
    if _sharepoint_configured():
        try:
            # The SharePoint list tracks every request regardless of storage
            # mode (see _sharepoint_configured()'s docstring) - a visibility
            # mirror for business users, not the backend's source of truth
            # (Table Storage is), so a failure here must not block submission.
            entity["sharepointListItemId"] = sharepoint_create_list_item(request_id, fields)
        except Exception as exc:  # noqa: BLE001
            logger.error(json.dumps({"event": "sharepoint_list_create_error", "requestId": request_id, "error": str(exc)}))
    client.upsert_entity(entity)


def update_tracking_status(request_id: str, status: str, **extra: Any) -> None:
    client = _table_client(TRACKING_TABLE_NAME)
    entity = {"PartitionKey": "request", "RowKey": request_id, "status": status, "updatedAt": datetime.now(timezone.utc).isoformat()}
    entity.update(extra)
    client.update_entity(entity, mode="merge")

    if _sharepoint_configured():
        try:
            record = get_tracking_record(request_id)
            item_id = record.get("sharepointListItemId", "") if record else ""
            if item_id:
                sharepoint_update_list_item(item_id, status=status, download_url=extra.get("downloadUrl"))
        except Exception as exc:  # noqa: BLE001 - mirror only, must not block the core status update
            logger.error(json.dumps({"event": "sharepoint_list_update_error", "requestId": request_id, "error": str(exc)}))


def get_tracking_record(request_id: str) -> dict[str, Any] | None:
    client = _table_client(TRACKING_TABLE_NAME)
    try:
        entity = client.get_entity("request", request_id)
    except ResourceNotFoundError:  # no row yet, or table not yet created - both 404
        return None
    return dict(entity)


def search_tracking(query: str) -> list[dict[str, Any]]:
    """Search by request ID (exact/prefix) or customer name (substring,
    case-insensitive). Searches both Table Storage (the backend's own fast
    tracking store, always populated for every request) and the SharePoint
    list (also always written to, independent of storage mode - see
    _sharepoint_configured()), merging by request ID so the list is proven to
    be a genuinely queryable tracking surface, not just a write-only mirror.
    Table Storage has no native substring filter, so customer-name matching is
    done client-side after a partition scan — fine at CEO-demo volume, not
    meant to scale."""
    client = _table_client(TRACKING_TABLE_NAME)
    query_upper = query.strip().upper()
    query_lower = query.strip().lower()
    results = []
    try:
        # query_entities() is lazy (returns an ItemPaged) - the service call, and
        # any ResourceNotFoundError for a not-yet-created table, happens on the
        # first iteration below, not on this call - so the loop must be inside
        # the try, not just the call.
        for entity in client.query_entities("PartitionKey eq 'request'"):
            row = dict(entity)
            if query_upper in row.get("RowKey", "").upper() or query_lower in row.get("customerName", "").lower():
                results.append(row)
    except ResourceNotFoundError:  # table not created yet - no requests submitted at all
        pass

    if _sharepoint_configured():
        seen_ids = {r.get("RowKey", "") for r in results}
        for sp_row in sharepoint_search_list(query):
            if sp_row["requestId"] not in seen_ids:
                results.append(sp_row)

    results.sort(key=lambda r: r.get("createdAt", ""), reverse=True)
    return results[:25]


# ---------------------------------------------------------------------------
# Document storage — the SharePoint/Blob boolean toggle.
# ---------------------------------------------------------------------------
def issue_upload_sas(session_id: str, file_name: str) -> dict[str, str]:
    """Short-lived, create+write-only user-delegation SAS so frontend.py can PUT a
    file straight to Blob Storage — file bytes never pass through this app. Only
    used when USE_SHAREPOINT is False; ported from project-2's CommonHandler.py."""
    now = datetime.now(timezone.utc)
    starts_on = now - timedelta(minutes=5)
    expires_on = now + timedelta(minutes=UPLOAD_SAS_TTL_MINUTES)

    blob_path = f"pending/{session_id}/{file_name}"
    client = _blob_service_client()
    blob_client = client.get_blob_client(container=STORAGE_UPLOADS_CONTAINER, blob=blob_path)

    user_delegation_key = client.get_user_delegation_key(starts_on, expires_on)
    sas_token = generate_blob_sas(
        account_name=client.account_name,
        container_name=STORAGE_UPLOADS_CONTAINER,
        blob_name=blob_path,
        user_delegation_key=user_delegation_key,
        permission=BlobSasPermissions(create=True, write=True),
        start=starts_on,
        expiry=expires_on,
    )
    return {
        "sasUrl": f"{blob_client.url}?{sas_token}",
        "blobPath": blob_path,
        "expiresOn": expires_on.isoformat(),
        "storageMode": "blob",
    }


def _graph_token() -> str:
    """SharePoint-only: the site lives in a different Entra tenant than this
    Function App's UAMI, which cannot cross tenant boundaries - so this uses
    the tenant-scoped app registration (client credentials flow) instead of
    managed identity. Raises clearly if the SharePoint app settings aren't
    configured, rather than falling back to a token that would silently fail
    against the wrong tenant."""
    if not (SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET):
        raise RuntimeError(
            "USE_SHAREPOINT is true but SHAREPOINT_TENANT_ID/SHAREPOINT_CLIENT_ID/"
            "SHAREPOINT_CLIENT_SECRET are not fully configured."
        )
    credential = ClientSecretCredential(
        tenant_id=SHAREPOINT_TENANT_ID, client_id=SHAREPOINT_CLIENT_ID, client_secret=SHAREPOINT_CLIENT_SECRET
    )
    return credential.get_token(GRAPH_SCOPE).token


def issue_sharepoint_upload_session(session_id: str, file_name: str) -> dict[str, str]:
    """SharePoint equivalent of issue_upload_sas() — a Graph upload session the
    browser can PUT bytes to directly (same 'bytes never touch this app' shape).
    Verified end-to-end against a real SharePoint site. Unlike a Blob SAS PUT, Graph's upload session requires a
    Content-Range header on the PUT even for a single-shot upload - frontend.py
    reads the storageMode field below to decide whether to send it (confirmed
    via a live HTTP 400 when it was missing - fixed, not assumed)."""
    if not SHAREPOINT_SITE_URL or not SHAREPOINT_DRIVE_ID:
        raise RuntimeError(
            "USE_SHAREPOINT is true but SHAREPOINT_SITE_URL/SHAREPOINT_DRIVE_ID are not configured. "
            "No SharePoint site has been provisioned for this project yet — see project memory."
        )
    token = _graph_token()
    path = f"pending/{session_id}/{file_name}"
    response = requests.post(
        f"https://graph.microsoft.com/v1.0/drives/{SHAREPOINT_DRIVE_ID}/root:/{path}:/createUploadSession",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30,
    )
    response.raise_for_status()
    body = response.json()
    return {
        "sasUrl": body["uploadUrl"],
        "blobPath": path,
        "expiresOn": body.get("expirationDateTime", ""),
        "storageMode": "sharepoint",
    }


def _sharepoint_download(path: str) -> bytes:
    """Reads a file's bytes straight out of the SharePoint drive via Graph."""
    token = _graph_token()
    response = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{SHAREPOINT_DRIVE_ID}/root:/{path}:/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=60,
    )
    response.raise_for_status()
    return response.content


def _stage_sharepoint_upload_for_processing(blob_path: str) -> None:
    """generate_otc (the existing, unmodified backend) can only read source
    documents from Blob Storage - its own download_uploaded_file() calls
    blob_client.download_blob() directly, with no SharePoint awareness
    (CONFIRMED via a live Application Insights traceback: ResourceNotFoundError/
    BlobNotFound when generate_otc was handed a path that only existed in
    SharePoint). So when USE_SHAREPOINT is true, this stages a transient copy
    of each uploaded document into Blob Storage at the same relative path,
    purely so generate_otc's existing contract keeps working unmodified.
    SharePoint remains the system of record the user/tab sees; Blob Storage
    here is just the AI pipeline's working cache."""
    content = _sharepoint_download(blob_path)
    client = _blob_service_client()
    client.get_blob_client(container=STORAGE_UPLOADS_CONTAINER, blob=blob_path).upload_blob(content, overwrite=True)


_sharepoint_list_id_cache: str | None = None


def _sharepoint_ensure_list() -> str:
    """Finds (or creates) the SharePoint *list* that tracks requests — the "own
    SharePoint list" the original spec asked for, distinct from the document
    library this project already writes files to (a List in the Lists app, not
    a folder of files). Looked up by display name and auto-created with a
    matching column set on first use so no manual SharePoint setup is needed
    beyond the site itself. Cached in-process; re-discovered on cold start."""
    global _sharepoint_list_id_cache
    if _sharepoint_list_id_cache:
        return _sharepoint_list_id_cache
    if not SHAREPOINT_SITE_ID:
        raise RuntimeError(
            "USE_SHAREPOINT is true but SHAREPOINT_SITE_ID is not configured "
            "(needed for the Lists API, distinct from SHAREPOINT_DRIVE_ID)."
        )
    token = _graph_token()
    headers = {"Authorization": f"Bearer {token}"}

    lookup = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists",
        headers=headers,
        params={"$filter": f"displayName eq '{SHAREPOINT_LIST_NAME}'"},
        timeout=30,
    )
    lookup.raise_for_status()
    existing = lookup.json().get("value", [])
    if existing:
        _sharepoint_list_id_cache = existing[0]["id"]
        return _sharepoint_list_id_cache

    create = requests.post(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists",
        headers={**headers, "Content-Type": "application/json"},
        json={
            "displayName": SHAREPOINT_LIST_NAME,
            "list": {"template": "genericList"},
            "columns": [
                {"name": "RequestId", "text": {}},
                {"name": "CustomerName", "text": {}},
                {"name": "RequestType", "text": {}},
                {"name": "FacilityType", "text": {}},
                {"name": "Status", "text": {}},
                {"name": "DownloadUrl", "text": {}},
            ],
        },
        timeout=30,
    )
    create.raise_for_status()
    _sharepoint_list_id_cache = create.json()["id"]
    return _sharepoint_list_id_cache


def sharepoint_create_list_item(request_id: str, fields: dict[str, str]) -> str:
    """Creates one row in the SharePoint tracking list for a new request.
    Returns the new item's id, which callers store (tracking table's
    sharepointListItemId field) so later status updates can PATCH it directly
    instead of relying on Graph's unreliable custom-column filtering."""
    list_id = _sharepoint_ensure_list()
    token = _graph_token()
    response = requests.post(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists/{list_id}/items",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "fields": {
                "Title": request_id,
                "RequestId": request_id,
                "CustomerName": fields.get("customerName", ""),
                "RequestType": fields.get("requestType", ""),
                "FacilityType": fields.get("facilityType", ""),
                "Status": ConstMessages.STATUS_CREATED,
            }
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["id"]


def sharepoint_update_list_item(item_id: str, status: str | None = None, download_url: str | None = None) -> None:
    """Patches an existing SharePoint list item's fields as a request's status
    changes - mirrors update_tracking_status()'s Table Storage write."""
    body: dict[str, str] = {}
    if status is not None:
        body["Status"] = status
    if download_url is not None:
        body["DownloadUrl"] = download_url
    if not body:
        return
    list_id = _sharepoint_ensure_list()
    token = _graph_token()
    response = requests.patch(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists/{list_id}/items/{item_id}/fields",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
        timeout=30,
    )
    response.raise_for_status()


def sharepoint_search_list(query: str) -> list[dict[str, Any]]:
    """Searches the SharePoint tracking list directly via Graph, independent of
    Table Storage - called from search_tracking() so the list is proven to be
    a genuinely queryable tracking surface, not just a write-only mirror.
    Client-side substring filter after a full list scan, same reasoning as
    search_tracking()'s own Table Storage fallback: Graph's list-item $filter
    on custom (non-indexed) columns is unreliable. Degrades to an empty list
    on any failure rather than breaking the whole /api/search response."""
    try:
        list_id = _sharepoint_ensure_list()
        token = _graph_token()
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists/{list_id}/items",
            headers={"Authorization": f"Bearer {token}"},
            params={"$expand": "fields", "$top": "200"},
            timeout=30,
        )
        response.raise_for_status()
    except Exception as exc:  # noqa: BLE001 - search must degrade gracefully, not 500 the endpoint
        logger.error(json.dumps({"event": "sharepoint_list_search_error", "error": str(exc)}))
        return []

    query_upper = query.strip().upper()
    query_lower = query.strip().lower()
    results = []
    for item in response.json().get("value", []):
        item_fields = item.get("fields", {})
        request_id = item_fields.get("RequestId", "")
        customer_name = item_fields.get("CustomerName", "")
        if query_upper in request_id.upper() or query_lower in customer_name.lower():
            results.append(
                {
                    "requestId": request_id,
                    "RowKey": request_id,
                    "customerName": customer_name,
                    "requestType": item_fields.get("RequestType", ""),
                    "facilityType": item_fields.get("FacilityType", ""),
                    "status": item_fields.get("Status", ""),
                    "downloadUrl": item_fields.get("DownloadUrl", ""),
                    "createdAt": "",
                    "source": "sharepoint",
                }
            )
    return results


def issue_upload_session(session_id: str, file_name: str) -> dict[str, str]:
    """The one entry point frontend.py's /api/sas route calls — branches on the
    boolean toggle so nothing above this line needs to know which backend is active."""
    return issue_sharepoint_upload_session(session_id, file_name) if USE_SHAREPOINT else issue_upload_sas(session_id, file_name)


def upload_generated_document(content: bytes, file_name: str) -> str:
    """Uploads the generated .docx (received from generate_otc) to whichever store
    the toggle selects, returns a URL frontend.py can show as the Download link."""
    if USE_SHAREPOINT:
        if not SHAREPOINT_SITE_URL or not SHAREPOINT_DRIVE_ID:
            raise RuntimeError("USE_SHAREPOINT is true but SharePoint is not provisioned yet (see project memory).")
        token = _graph_token()
        path = f"generated/{file_name}"
        response = requests.put(
            f"https://graph.microsoft.com/v1.0/drives/{SHAREPOINT_DRIVE_ID}/root:/{path}:/content",
            headers={"Authorization": f"Bearer {token}"},
            data=content,
            timeout=60,
        )
        response.raise_for_status()
        return response.json().get("webUrl", "")

    client = _blob_service_client()
    blob_client = client.get_blob_client(container=STORAGE_GENERATED_CONTAINER, blob=file_name)
    content_type = Defaults.CONTENT_TYPES.get(
        file_name.rsplit(".", 1)[-1].lower() if "." in file_name else "", "application/octet-stream"
    )
    blob_client.upload_blob(content, overwrite=True, content_settings=ContentSettings(content_type=content_type))
    sas_token = generate_blob_sas(
        account_name=client.account_name,
        container_name=STORAGE_GENERATED_CONTAINER,
        blob_name=file_name,
        account_key=None,
        user_delegation_key=client.get_user_delegation_key(
            datetime.now(timezone.utc) - timedelta(minutes=5), datetime.now(timezone.utc) + timedelta(days=7)
        ),
        permission=BlobSasPermissions(read=True),
        expiry=datetime.now(timezone.utc) + timedelta(days=7),
    )
    return f"{blob_client.url}?{sas_token}"


# ---------------------------------------------------------------------------
# Queue hand-off — Storage Queue (not Service Bus), per instruction: avoid
# Service Bus unless there's a concrete reason it's required. A Storage Queue
# on the existing storage account needs no new resource type.
# ---------------------------------------------------------------------------
def enqueue_request(request_id: str, payload: dict[str, Any]) -> None:
    """host.json declares QueuesOptions.MessageEncoding "Base64" (the WebJobs Queue
    trigger extension's own default) — the message body MUST be base64-encoded or
    the trigger fails to decode it at the binding level, before process_queue_message
    is ever invoked (no exception, no "Executing" trace — just silent, instant
    poisoning). azure-storage-queue's send_message() does NOT base64-encode by
    default, unlike the .NET SDK conventions WebJobs assumes, so this has to be
    done explicitly here."""
    import base64

    client = _queue_client()
    try:
        client.create_queue()
    except ResourceExistsError:
        pass
    body = json.dumps({"requestId": request_id, **payload}, default=str)
    client.send_message(base64.b64encode(body.encode("utf-8")).decode("ascii"))


# ---------------------------------------------------------------------------
# generate_otc — ported byte-for-byte (URL shape + auth header) from
# project-2-best-practice-solution/CommonHandler.py::trigger_otc_generation(),
# which already calls this exact APIM instance successfully. Payload fields
# beyond what's shown are UNCONFIRMED — see module docstring.
# ---------------------------------------------------------------------------
async def call_generate_otc(payload: dict[str, Any]) -> dict[str, Any]:
    if not APIM_BASE_URL:
        raise RuntimeError("APIM_BASE_URL environment variable not set.")

    headers = {"Content-Type": "application/json"}
    if APIM_SUBSCRIPTION_KEY:
        headers["Ocp-Apim-Subscription-Key"] = APIM_SUBSCRIPTION_KEY

    async with aiohttp.ClientSession() as session:
        async with session.post(
            # External APIM path is the API's own path ("otc") + the operation's
            # urlTemplate ("/otc/generate") = "/otc/otc/generate" — confirmed by
            # reading the actual operation definition (az rest GET .../apis/
            # otc-orchestrator/operations), not assumed. The backend-facing path
            # (serviceUrl + urlTemplate, no API-path prefix) is plain
            # ".../api/v1/otc/generate", matching generate_otc's own InvokeUrlTemplate.
            f"{APIM_BASE_URL}/otc/otc/generate",
            headers=headers,
            json=payload,
            timeout=aiohttp.ClientTimeout(total=Defaults.GENERATE_OTC_TIMEOUT_SECONDS),
        ) as response:
            body_text = await response.text()
            if response.status not in (200, 202):
                logger.error(json.dumps({"event": "generate_otc_error", "status": response.status, "body": body_text[:500]}))
                raise RuntimeError(f"generate_otc failed: HTTP {response.status}")
            return json.loads(body_text)


# ---------------------------------------------------------------------------
# Proactive messaging — direct Bot Connector REST call with an MSI-acquired
# token, ported from project-2's az-function-apps/otc-orchestrator/ProactiveBot.py
# (there is no Python binding for the SDK's server-to-server Connector client).
# ---------------------------------------------------------------------------
def _bot_connector_token() -> str:
    credential = ManagedIdentityCredential(client_id=MICROSOFT_APP_ID) if MICROSOFT_APP_ID else DefaultAzureCredential()
    return credential.get_token(BOT_CONNECTOR_SCOPE).token


def send_proactive_card(conversation_reference: dict[str, Any], card_content: dict[str, Any]) -> None:
    service_url = conversation_reference["serviceUrl"]
    conversation_id = conversation_reference["conversation"]["id"]
    activity = {
        "type": "message",
        "id": str(uuid.uuid4()),
        "serviceUrl": service_url,
        "channelId": conversation_reference["channelId"],
        "conversation": conversation_reference["conversation"],
        "from": conversation_reference.get("agent") or conversation_reference.get("bot") or {"id": MICROSOFT_APP_ID},
        "recipient": conversation_reference.get("user"),
        "locale": conversation_reference.get("locale"),
        "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive", "content": card_content}],
    }
    token = _bot_connector_token()
    response = requests.post(
        f"{service_url.rstrip('/')}/v3/conversations/{conversation_id}/activities",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        data=json.dumps(activity, default=str),
        timeout=30,
    )
    response.raise_for_status()


def build_result_card(tracking: dict[str, Any]) -> dict[str, Any]:
    facts = [
        {"title": "Customer Name", "value": tracking.get("customerName", "")},
        {"title": "Request Type", "value": tracking.get("requestType", "")},
        {"title": "Facility Type", "value": tracking.get("facilityType", "")},
        {"title": "Request ID", "value": tracking.get("requestId", "")},
        {"title": "Completed", "value": tracking.get("updatedAt", "")},
    ]
    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {"type": "TextBlock", "text": "✔ Draft OTC Generated Successfully", "weight": "Bolder", "size": "Large", "color": "Good", "wrap": True},
            {"type": "FactSet", "facts": facts},
            {"type": "TextBlock", "text": f"{tracking.get('fileName', '')} ({tracking.get('fileSizeKb', 0)} KB)", "wrap": True, "isSubtle": True},
        ],
        "actions": [{"type": "Action.OpenUrl", "title": "Download Draft OTC", "url": tracking.get("downloadUrl", "")}],
    }


# ---------------------------------------------------------------------------
# The full pipeline a queue message triggers — see function_app.py's queue
# trigger. Centralized here so function_app.py stays pure hosting plumbing.
# ---------------------------------------------------------------------------
def _build_generate_otc_payload(message: dict[str, Any]) -> dict[str, Any]:
    """Translate this app's own submit payload shape (documents.investment_memo.path /
    documents.financial_model.path) into generate_otc's actual expected shape.

    CONFIRMED (not assumed) via a live Application Insights traceback from
    generate_otc's own code (func-otc-agent-dev-0efu4i, requests/exceptions
    tables, 2026-11-10): it reads payload["investmentMemoBlobPath"] directly
    (KeyError on that exact name when absent) - a flat field, not the nested
    documents.investment_memo.path shape this app originally sent. The
    financial-model counterpart name is inferred by the same naming
    convention (not yet confirmed by its own traceback) - if that guess is
    also wrong, generate_otc will raise the same KeyError pattern for it,
    which will show up the same way and should be fixed the same way.
    """
    documents = message.get("documents") or {}
    payload = dict(message)
    payload["investmentMemoBlobPath"] = (documents.get("investment_memo") or {}).get("path", "")
    payload["financialModelBlobPath"] = (documents.get("financial_model") or {}).get("path", "")
    return payload


async def process_request(message: dict[str, Any]) -> None:
    request_id = message["requestId"]
    try:
        update_tracking_status(request_id, ConstMessages.STATUS_IN_REVIEW)
        update_tracking_status(request_id, ConstMessages.STATUS_PROCESSING)

        if USE_SHAREPOINT:
            # generate_otc can only read from Blob Storage (see
            # _stage_sharepoint_upload_for_processing's docstring) - stage
            # transient copies before calling it.
            documents = message.get("documents") or {}
            for doc in (documents.get("investment_memo"), documents.get("financial_model")):
                if doc and doc.get("path"):
                    _stage_sharepoint_upload_for_processing(doc["path"])

        result = await call_generate_otc(_build_generate_otc_payload(message))

        # generate_otc's exact response shape is UNCONFIRMED (no access to its
        # source — see module docstring). Rather than assume one naming
        # convention, try every plausible key for each field (snake_case,
        # camelCase, and a couple of likely alternates) and log clearly if none
        # match, so a contract mismatch is immediately diagnosable instead of
        # silently producing an empty download link.
        inner = result.get("result") if isinstance(result.get("result"), dict) else result

        def _first(*keys: str) -> Any:
            for key in keys:
                if inner.get(key) not in (None, ""):
                    return inner[key]
            return None

        file_name = _first("file_name", "fileName", "filename", "name") or f"{request_id}.docx"
        content_b64 = _first("content_base64", "contentBase64", "content", "fileContent", "base64")
        download_url_field = _first("download_url", "downloadUrl", "documentUrl", "document_url", "url", "fileUrl")
        file_size_kb_field = _first("file_size_kb", "fileSizeKb", "size_kb", "sizeKb")

        if content_b64:
            import base64

            content = base64.b64decode(content_b64)
            download_url = upload_generated_document(content, file_name)
            file_size_kb = round(len(content) / 1024)
        elif download_url_field:
            if USE_SHAREPOINT:
                # generate_otc always places its own output in Blob Storage
                # directly (CONFIRMED - its own fixed behavior, not
                # toggle-aware) regardless of this app's setting, so mirror
                # that output into SharePoint here to keep SharePoint the
                # system of record the user/tab sees when the toggle is on.
                content = requests.get(download_url_field, timeout=60).content
                download_url = upload_generated_document(content, file_name)
                file_size_kb = round(len(content) / 1024)
            else:
                # generate_otc already placed the file in storage itself and
                # just returned a URL.
                download_url = download_url_field
                file_size_kb = file_size_kb_field or 0
        else:
            logger.error(
                json.dumps(
                    {
                        "event": "generate_otc_unrecognized_response_shape",
                        "requestId": request_id,
                        "responseKeys": list(inner.keys()) if isinstance(inner, dict) else type(inner).__name__,
                    }
                )
            )
            raise RuntimeError(
                "generate_otc returned a response with no recognized content/download-url field "
                f"(keys seen: {list(inner.keys()) if isinstance(inner, dict) else inner!r})"
            )

        update_tracking_status(
            request_id,
            ConstMessages.STATUS_GENERATED,
            downloadUrl=download_url,
            fileName=file_name,
            fileSizeKb=file_size_kb,
        )
        update_tracking_status(request_id, ConstMessages.STATUS_COMPLETED)

        conversation_reference = get_request_conversation_reference(request_id)
        if conversation_reference:
            tracking = get_tracking_record(request_id) or {}
            send_proactive_card(conversation_reference, build_result_card(tracking))

    except Exception as exc:  # noqa: BLE001 - must not crash the queue trigger
        logger.error(json.dumps({"event": "process_request_error", "requestId": request_id, "error": str(exc)}))
        update_tracking_status(request_id, ConstMessages.STATUS_FAILED, errorMessage=str(exc))


# ---------------------------------------------------------------------------
# Bot handlers — minimal. The real UI is frontend.py's tab; the bot's only job
# is the welcome message (with a deep link to the tab) and remembering each
# user's conversationReference for proactive messaging later.
# ---------------------------------------------------------------------------
AGENT_APP = AgentApplication[TurnState](storage=STORAGE, adapter=ADAPTER, authorization=AUTHORIZATION, **agents_sdk_config)


def _remember_user(context: TurnContext) -> None:
    user_id = context.activity.from_property.id if context.activity.from_property else None
    if not user_id:
        return
    conversation_reference = context.activity.get_conversation_reference()
    save_user_conversation_reference(user_id, conversation_reference.model_dump(mode="json", by_alias=True))


@AGENT_APP.conversation_update("membersAdded")
async def on_members_added(context: TurnContext, _state: TurnState):
    _remember_user(context)
    await context.send_activity(
        "Welcome to the OTC Assistant. Open the **New OTC Request** tab above to submit a request and track it live."
    )
    return True


@AGENT_APP.activity("message")
async def on_message(context: TurnContext, _state: TurnState):
    _remember_user(context)
    await context.send_activity("Open the **New OTC Request** tab above to submit or track a request.")


@AGENT_APP.error
async def on_error(context: TurnContext, error: Exception):
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    await context.send_activity("The bot encountered an error.")
