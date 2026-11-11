"""Azure Functions entrypoint — pure hosting plumbing. Every real decision (storage
toggle, queue shape, generate_otc call, proactive messaging) lives in agent.py;
this file only wires HTTP/queue triggers to it. Same reasoning as every other
project here: this subscription has zero App Service quota (reconfirmed before
this project started), so Azure Functions Consumption is the hosting surface.
"""

import asyncio
import json
import logging

import azure.functions as func

import agent
import frontend
from Constants import EnvVariables, Defaults, REQUIRED_FIELDS

ALLOWED_UPLOAD_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx"}

logger = logging.getLogger("teams_bot_final_product.function_app")
app = func.FunctionApp()


class AzureFunctionsRequestAdapter:
    """Implements microsoft_agents.hosting.core.http.HttpRequestProtocol over a
    func.HttpRequest — identical in role to every other project's copy of this class."""

    def __init__(self, request: func.HttpRequest):
        self._request = request

    @property
    def method(self) -> str:
        return self._request.method

    @property
    def headers(self):
        return self._request.headers

    async def json(self):
        return self._request.get_json()

    def get_claims_identity(self):
        return None

    def get_path_param(self, name: str) -> str:
        return self._request.route_params.get(name)


@app.route(route="messages", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
async def messages(req: func.HttpRequest) -> func.HttpResponse:
    adapted_request = AzureFunctionsRequestAdapter(req)
    try:
        http_response = await agent.ADAPTER.process_request(adapted_request, agent.AGENT_APP)
    except Exception as exc:  # noqa: BLE001 - malformed/unexpected activity must not 500 to the Bot Connector
        logger.error(json.dumps({"event": "messages_route_error", "error": str(exc)}))
        return func.HttpResponse(json.dumps({"error": "unable to process activity"}), status_code=200, mimetype="application/json")
    return func.HttpResponse(
        body=json.dumps(http_response.body) if http_response.body is not None else "",
        status_code=http_response.status_code,
        headers=http_response.headers or {},
        mimetype="application/json",
    )


@app.route(route="form", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def form(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse(frontend.render_page(), mimetype="text/html")


@app.route(route="sas", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def sas(req: func.HttpRequest) -> func.HttpResponse:
    session_id = req.params.get("sessionId", "")
    file_name = req.params.get("fileName", "")
    if not session_id or not file_name:
        return func.HttpResponse(json.dumps({"error": "sessionId and fileName are required"}), status_code=400, mimetype="application/json")
    extension = "." + file_name.rsplit(".", 1)[-1].lower() if "." in file_name else ""
    if extension not in ALLOWED_UPLOAD_EXTENSIONS:
        # Enforced server-side, not just by the file picker's `accept` attribute in
        # frontend.py — that's trivially bypassed (drag-and-drop, devtools), and the
        # user's requirement is "PDF/Word/Excel only, other formats rejected".
        return func.HttpResponse(
            json.dumps({"error": f"Unsupported file type '{extension or file_name}'. Allowed: PDF, DOC/DOCX, XLS/XLSX."}),
            status_code=400,
            mimetype="application/json",
        )
    try:
        result = agent.issue_upload_session(session_id, file_name)
    except Exception as exc:  # noqa: BLE001
        logger.error(json.dumps({"event": "sas_error", "error": str(exc)}))
        return func.HttpResponse(json.dumps({"error": str(exc)}), status_code=502, mimetype="application/json")
    return func.HttpResponse(json.dumps(result), mimetype="application/json")


@app.route(route="submit", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def submit(req: func.HttpRequest) -> func.HttpResponse:
    try:
        payload = req.get_json()
    except ValueError:
        return func.HttpResponse(json.dumps({"error": "Invalid JSON"}), status_code=400, mimetype="application/json")

    missing = [f for f in REQUIRED_FIELDS if not payload.get(f)]
    if missing:
        return func.HttpResponse(json.dumps({"error": f"Missing required fields: {', '.join(missing)}"}), status_code=400, mimetype="application/json")

    request_id = agent.build_request_id()
    requested_by = payload.get("requestedBy", "")
    teams_user_id = payload.get("teamsUserId", "")

    agent.create_tracking_record(request_id, payload, requested_by, teams_user_id)

    conversation_reference = agent.get_user_conversation_reference(teams_user_id) if teams_user_id else None
    if conversation_reference:
        agent.save_request_conversation_reference(request_id, conversation_reference)

    agent.enqueue_request(request_id, payload)

    return func.HttpResponse(json.dumps({"requestId": request_id, "status": "submitted"}), mimetype="application/json")


@app.route(route="status", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def status(req: func.HttpRequest) -> func.HttpResponse:
    request_id = req.params.get("id", "")
    if not request_id:
        return func.HttpResponse(json.dumps({"error": "id is required"}), status_code=400, mimetype="application/json")
    record = agent.get_tracking_record(request_id)
    if not record:
        return func.HttpResponse(json.dumps({"error": "not found"}), status_code=404, mimetype="application/json")
    # frontend.py's pollStatus() reads data.record, not the bare body - keep this
    # wrapped (same convention as /api/search's {"results": [...]}), not a bare object.
    return func.HttpResponse(json.dumps({"record": record}, default=str), mimetype="application/json")


@app.route(route="search", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def search(req: func.HttpRequest) -> func.HttpResponse:
    query = req.params.get("q", "")
    results = agent.search_tracking(query) if query else []
    return func.HttpResponse(json.dumps({"results": results}, default=str), mimetype="application/json")


@app.queue_trigger(arg_name="msg", queue_name=Defaults.STORAGE_QUEUE_NAME, connection="StorageConnection")
# IMPORTANT: "StorageConnection" is an identity-based connection, deliberately NOT
# "AzureWebJobsStorage" — the Functions runtime storage account
# is a *different* account from this project's business-data storage
# (STORAGE_ACCOUNT_NAME), which is where agent.enqueue_request()
# actually sends messages via _queue_client(). Using AzureWebJobsStorage here would
# point the trigger at the wrong account's queue and nothing would ever fire.
# Requires these app settings (identity-based connection, UAMI, no connection string):
#   StorageConnection__accountName = <STORAGE_ACCOUNT_NAME>
#   StorageConnection__credential  = managedidentity
#   StorageConnection__clientId    = <AZURE_CLIENT_ID>  (the UAMI's client id)
def process_queue_message(msg: func.QueueMessage) -> None:
    body = json.loads(msg.get_body().decode("utf-8"))
    asyncio.run(agent.process_request(body))
