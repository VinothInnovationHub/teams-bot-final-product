"""The one piece of plumbing agent.py cannot avoid under Azure Functions hosting.

Some teams' own agent.py constructs `CloudAdapter` from `microsoft_agents.hosting.aiohttp`
— hard-wired to aiohttp's Request/Response types, since that reference runs as a
raw aiohttp server (main.py + start_server.py). This subscription has zero App
Service quota on every SKU/region (verified in Project 1), so there is nowhere to
run that aiohttp server; Azure Functions (Consumption) is the only available free
surface for a persistent custom bot endpoint. `HttpAdapterBase` already contains
every bit of the actual request-processing logic (JWT claims, activity validation,
invoke handling) — this class adds nothing beyond it, the same way
`microsoft_agents.hosting.aiohttp.CloudAdapter` itself is just a thin aiohttp-specific
subclass of the same base. See Project 1/2's identical class for the precedent.
"""

from microsoft_agents.hosting.core.http import HttpAdapterBase


class AzureFunctionsAdapter(HttpAdapterBase):
    pass
