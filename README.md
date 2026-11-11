# teams-bot-final-product

A Microsoft Teams bot + tab for submitting an "Offer-to-Close" style request:
a user fills a form, uploads two documents, an AI pipeline drafts the output
document, and the user gets a live progress tracker plus a Teams notification
when it's done.

Hosted entirely on **Azure Functions** (serverless) — one Function App runs
both the web-facing routes (form, upload-link issuance, submit, status,
search) and the background processing (queue trigger -> AI call -> storage ->
notification).

## Status

Proven and verified end-to-end: real uploads, real AI processing via APIM,
real generated documents, real Teams notifications, in both supported
storage modes (Blob and SharePoint).

## The request flow, step by step

1. User opens the tab -> sees the form
2. User fills the form, picks 2 files, clicks Submit
3. Browser asks the server for a short-lived upload link (a Blob SAS token,
   or a Graph upload session for SharePoint)
4. Browser uploads the 2 files **directly** to storage — this app's own code
   never touches the file bytes
5. Browser posts the form data to `/api/submit`
6. Server creates a tracking record (status = "Request Created"), and a
   matching SharePoint list row if SharePoint is configured
7. Server drops a message on the Storage Queue
8. Server replies immediately with the request ID — the browser starts
   polling `/api/status` every few seconds

*(the rest happens in the background, independent of the browser)*

9. The queue trigger wakes the Function App when the message arrives
10. It calls the AI backend **through APIM** (subscription-key gated)
11. The AI backend reads the files, runs Document Intelligence + an LLM,
    returns a result
12. The app uploads the generated document to whichever store is active
13. The tracking record is updated to "Completed" with a real download link
14. A proactive message is sent into the Teams chat: "your draft is ready"
15. The browser's next poll sees "Completed" and shows the result

## Why there's no request-timeout risk

Steps 1-8 (the part the browser is waiting on) return a response in
milliseconds. Steps 9-15 (the slow AI work) run entirely in the background,
triggered by a queue message — nothing the user does is ever connected,
end to end, to the slow part.

## Security model

- Every call to the AI backend is gated by APIM (subscription key).
- The app authenticates to Azure using a managed identity — no long-lived
  secret is held in code or in application settings for Azure-to-Azure
  calls.
- The one thing an end user ever touches directly — the upload link — is a
  short-lived, purpose-built token (SAS or Graph upload session) that
  expires and can't be reused. Storage is never made public.

## Folder structure

```
teams-bot-final-product/
├── app_services/msbotservice/src/
│   ├── function_app.py              # all route + queue-trigger bindings
│   ├── agent.py                     # storage, tracking, upload-session
│   │                                 issuance, SharePoint, bot wiring
│   ├── azure_functions_adapter.py   # the Azure-Functions-specific hosting
│   │                                 adapter (HttpAdapterBase subclass)
│   ├── frontend.py                  # the entire Teams tab UI (HTML/CSS/JS,
│   │                                 no build step)
│   ├── Constants.py                 # every config/env-var name in one place
│   ├── .env.example                 # every environment variable this app
│   │                                 needs, with placeholders
│   ├── host.json
│   └── requirements.txt
└── manifest/
    ├── manifest.json                # Teams app manifest (placeholders —
    │                                 fill in your own bot/app id and
    │                                 hostname before packaging)
    ├── color.png
    └── outline.png
```

## Setup

1. **Fill in `.env.example`** with your own Azure resource values (storage
   account, APIM endpoint + key, bot identity, SharePoint details if using
   it) and set them as Application Settings on your Function App.
2. **Fill in `manifest.json`'s placeholders** (`id`, `botId`, the
   `websiteUrl`/`contentUrl`/`validDomains` hostnames) with your own bot's
   app ID and your Function App's hostname, then zip `manifest.json` +
   `color.png` + `outline.png` together to produce a sideloadable Teams app
   package.
3. **Deploy the Function App** with your preferred tooling (`func azure
   functionapp publish`, CI/CD, etc.).
4. **Grant the Function App's identity** Storage Blob/Table/Queue Data
   Contributor on the storage account, and (if using SharePoint) Graph
   `Sites.Selected` scoped to the one SharePoint site.
5. **Test both storage modes** by toggling `USE_SHAREPOINT` and submitting a
   real request through the tab.
