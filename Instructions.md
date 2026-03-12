# ITGlue to ServiceNow Knowledge Base Converter

This document provides instructions on how to use the `ITGlue-To-ServiceNow.ps1` script to parse ITGlue exports and prepare them for import into ServiceNow.

## What It Does

The script automates the process of parsing an ITGlue Knowledge Base export and generating JSON payloads formatted for ServiceNow. Its core functionalities include:

- **Recursive Folder Scanning**: It recursively scans through the export folder hierarchy to find document folders (which begin with `DOC-`). It intelligently handles any arbitrary folder structures (like "How To Guides", "Networking", etc.) by diving into them until it finds the target document folders.
- **Knowledge Article Processing**: When it finds a Knowledge Article (an HTML file accompanied by a matching customer folder), it parses the HTML.
- **Image Embedding**: Any local images referenced within the HTML file are converted into base64 data URIs and embedded directly within the HTML. This ensures images aren't lost when importing to ServiceNow.
- **Attachment Handling**: It scans the `attachments` root folder for any files associated with the processed document ID, reads them, and packages them into the payload.
- **Attachment-Only Processing**: It can detect and process entries that are purely file attachments without substantial HTML content (e.g., 0kb HTML files).
- **JSON Payload Generation**: It bundles the extracted HTML content, attachments, and metadata (like document name and company name) into a JSON payload and sends it to the configured ServiceNow REST API endpoint (currently configured with a mock endpoint for testing purposes).
- **Oversized Document Handling**: If a single document (HTML + all attachments) exceeds the 5 MB batch limit on its own, it is skipped entirely and never sent — it would cause a 500 error if attempted. Instead it is logged to a per-company report and a machine-readable tracking file for future processing.

## Prerequisites

- **ServiceNow Plugin**: The following plugin must be installed and activated on your ServiceNow instance:
  - `com.glide.hub.action_step.jsonparser`

- **ServiceNow System Property**: The following system property must be set on the **global** scope of your ServiceNow instance:

  | Property | Type | Value |
  |---|---|---|
  | `com.glide.transform.json.max-partial-length` | Integer | `8388608` |

  This increases the maximum allowed JSON payload size to 8 MB, which is required to process large document batches. To set it, navigate to **System Properties** (`sys_properties.list`) in ServiceNow and create or update the record accordingly.

## How to Execute It

1. Open **PowerShell** on your Windows machine.
2. Navigate to the directory where you saved the `ITGlue-To-ServiceNow.ps1` script using the `cd` command. For example:
   ```powershell
   cd "C:\Users\cesar\OneDrive\Scripting\ITGlue KB Manager"
   ```
3. Run the script:
   ```powershell
   .\ITGlue-To-ServiceNow.ps1
   ```

## The Interactive Wizard
When you run the script, an interactive wizard will guide you through the necessary steps:

1. **Root Folder Location**: The script will ask: *"Is the script in the root folder for the ITGlue Documentation? (Yes/No)"*
   - Type `Yes` (or `Y`) if you placed the script directly beside the `documents` folder of your export.
   - Type `No` (or `N`) if the script is stored separately. It will then prompt you to paste the file path to the ITGlue export root folder.
2. **ServiceNow Endpoint URI**: You will be prompted to paste the destination API endpoint for ServiceNow. Your ServiceNow team will provide this (e.g., `https://teamascenddev.service-now.com/api/...`).
3. **Integration Credentials**: The script will ask for the ServiceNow Integration Username and Password. The password input is secure and will be masked as you type. It will use these credentials to authenticate API calls using Basic Auth.

> **Note**: If you receive a red error message about "Execution Policy" when trying to run the script, PowerShell is blocking scripts from running for security purposes. You can temporarily allow the script to run for your current window by typing this command first:
> `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`

## Oversized Document Reports

When the script encounters a document that exceeds the 5 MB limit on its own, it skips it and writes two files in the root folder:

### Per-company markdown report — `<CompanyName>.md`
Created (or appended to) in `$SourcePath` the first time an oversized document is found for that company. Contains a table of all skipped documents:

| Column | Description |
|---|---|
| Document Name | The name of the skipped Knowledge Article or attachment |
| Size | Estimated size in MB |
| Folder Path | Full path to the `DOC-` folder that was skipped |
| Timestamp | Date and time the skip was recorded |

### Oversized tracking file — `.migration-oversized.txt`
Located in `$SourcePath`. Lists one folder path per line for every skipped document, across all companies and all runs. This file is **not** deleted on completion — it accumulates over time and is intended to be used by a future run of the script once the 5 MB limit has been increased.

---

# ServiceNow Script Include: ITGlueImportDocuments

This document covers the ServiceNow-side component (`ITGlueImportDocuments`) that receives the JSON payloads sent by the PowerShell script and processes them into Knowledge Bases, Knowledge Articles, and attachments inside ServiceNow.

**File**: `x_wemop_knvl_itg_0/sys_script_include/ITGlueImportDocuments.script.js`

## What It Does

- Receives a JSON payload containing one or more documents grouped by company name.
- Looks up or creates a **Knowledge Base** per company, assigning the correct domain from `core_company`.
- Creates a **Knowledge Article** per document within the corresponding Knowledge Base, inheriting the same domain.
- Attaches any binary files (base64-encoded) to the created Knowledge Article.
- Publishes the article after all attachments are processed.
- Provides a **company lookup endpoint** to verify whether companies exist in ServiceNow before sending documents.

## Instantiation & Debug Mode

The Script Include accepts an optional `mode` parameter on instantiation:

```js
// All logs enabled
var importer = new ITGlueImportDocuments("debug");

// Logs suppressed (default / production)
var importer = new ITGlueImportDocuments();
```

When `"debug"` is passed, all `gs.info`, `gs.warn`, and `gs.error` log entries are written. Without it, logging is completely silent.

## Public Methods

### `importDocuments(request)`
The main entry point. Accepts a REST API request whose body contains:
```json
{
  "document": [
    {
      "companyname": "Acme Corp",
      "documentname": "Network Overview",
      "documentcontent": "<html>...</html>",
      "attachments": [
        { "filename": "diagram.pdf", "contenttype": "application/pdf", "content": "<base64>" }
      ]
    }
  ]
}
```
For each document it:
1. Resolves or creates the Knowledge Base (with domain).
2. Creates the Knowledge Article in `draft` state with the correct domain.
3. Writes all attachments to the article.
4. Sets the article to `published`.

### `checkCompany(request, companyName)` — Dual-mode
- **`checkCompany(request)`** — Bulk mode. Accepts a REST request with `{ "companyNames": ["Acme", "..."] }` and returns `{ companyNames: [true/false, ...] }` indicating which companies exist in `core_company`.
- **`checkCompany(null, companyName)`** — Internal single-lookup mode. Queries `core_company` by name and returns the `sys_domain` value of the matching record, or `null` if not found.

## Domain Assignment

When a new Knowledge Base is created, the script:
1. Calls `checkCompany(null, companyName)` to resolve the `sys_domain` from the matching `core_company` record.
2. Sets `sys_domain` on the newly created `kb_knowledge_base` record.
3. Returns both `kbSysId` and `domain` to the caller.

The Knowledge Article then has its `sys_domain` set to the same domain value returned from the Knowledge Base step, ensuring both records belong to the correct domain.

If the company is not found in `core_company`, the KB and article are still created without a domain assignment (a warning is logged in debug mode).
