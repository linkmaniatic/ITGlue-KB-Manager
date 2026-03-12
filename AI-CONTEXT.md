# AI Context: ITGlue-To-ServiceNow.ps1 (ITGlue to ServiceNow Payload Generator)

## System Metadata
- **Target File**: `ITGlue-To-ServiceNow.ps1`
- **Primary Language**: PowerShell
- **Core Purpose**: Batch parsing of ITGlue structured exports to construct and transmit JSON payloads to a ServiceNow Knowledge Base API endpoint via an interactive wizard.
- **System Dependencies**: Standard PowerShell environment (no external module requirements). The previous Microsoft Word COM dependency has been removed.

## Architecture & Execution Flow
1. **Interactive Wizard & Setup**:
   - Prompts user to specify if the script is in the root ITGlue Export folder. If not, captures the `$SourcePath`.
   - Prompts for `$ServiceNowEndpoint` URI.
   - Prompts securely for `$ServiceNowUser` and `$ServiceNowPass` (Integration Credentials).
2. **Pre-flight Connection Test**:
   - Constructs an `Authorization: Basic` header using the provided credentials.
   - Performs an initial `Invoke-RestMethod` `GET` request to verify the `$ServiceNowEndpoint` is reachable and credentials are valid before proceeding. Exits on failure.
3. **Recursive Folder Scan (`Invoke-FolderScan`)**:
   - The script iterates recursively through directories.
   - **Logic B (Root Mapping)**: If a folder does NOT match the `DOC-*` pattern, it is treated as the new `CompanyName` root, and scanning proceeds within it.
   - **Logic A (Knowledge Article)**: If a folder matches `DOC-{customer_id}-{document_id} {document_name}`, contains an HTML file, and contains a folder matching `{customer_id}`:
     - The HTML file is processed (local images base64 embedded).
     - Attachments are pulled from `attachments/documents/{document_id}`.
     - A JSON payload is constructed containing the HTML `documentcontent` and `attachments`.
   - **Logic C (Attachment Only)**: If a folder matches `DOC-*` but ends in a file extension (e.g. `.pdf`) and contains a 0-byte HTML file:
     - The folder is treated purely as a container for an attachment.
     - The HTML is skipped (left blank in payload).
     - The binary file is pulled from `attachments/documents/{document_id}` and mapped into the payload.
4. **Image Processing (`Convert-HtmlAndImages`)**:
   - Parses `src="([^"]+)"` paths in the HTML.
   - Converts local responsive images to Base64 `data:` URI streams inside the HTML payload by evaluating magic bytes signatures (PNG, JPEG, GIF, BMP).
5. **Attachment Processing (`Get-DocumentAttachments`)**:
   - Scans the `attachments` root folder parallel to the document ID.
   - Reads raw file bytes and encodes them as **Base64** strings for lossless binary transport inside the JSON payload. The ServiceNow receiver must decode Base64 to recover the original bytes.
6. **Oversized Document Guard (`Add-DocumentToBatch`)**:
   - Before queuing, `Get-DocumentSizeEstimate` is called on the document.
   - If the estimated size exceeds the 5 MB cap on its own, `Write-OversizedDocument` is called and the function returns immediately — the document is never added to any batch.
   - If the document is within the cap but adding it would overflow the current batch, the current batch is flushed first via `Send-CurrentBatch`, then the document is queued.
7. **Payload Dispatch (`Send-CurrentBatch`)**:
   - Documents are queued into a batch array via `Add-DocumentToBatch`. Size is estimated per document (string lengths + overhead).
   - When the batch would exceed **5 MB**, it is automatically flushed before adding the next document.
   - The JSON payload uses an **array** structure: `{ "document": [ {...}, {...} ] }`.
   - Serialized via `ConvertTo-Json -Depth 10 -Compress` and submitted via `Invoke-RestMethod` (`POST`) with `Authorization: Basic` header.
   - Any remaining documents are flushed after the scan completes.

## Key Methods / Functions
- `Invoke-FolderScan($FolderPath, $CompanyName)`: The recursive core engine interpreting folder logic based on pattern-matching conventions.
- `Get-DocumentAttachments($DocumentId)`: Returns an array of attachment objects consisting of `filename`, `contenttype`, and `content` (Base64-encoded bytes).
- `Convert-HtmlAndImages($HtmlFilePath)`: Responsible for parsing the HTML string and embedding local asset resources.
- `Get-DocumentSizeEstimate($Document)`: Estimates JSON byte size of a document hashtable without serialization.
- `Add-DocumentToBatch($Document, $FolderPath)`: Guards against oversized documents (> 5 MB) before queuing; logs them via `Write-OversizedDocument` and returns. Auto-flushes the current batch if adding the document would exceed the cap.
- `Write-OversizedDocument($CompanyName, $DocumentName, $FolderPath, $DocumentSizeBytes)`: Writes to two outputs — the per-company `<CompanyName>.md` markdown report (created on first entry, appended thereafter) and the `.migration-oversized.txt` tracking file (one folder path per line).
- `Send-CurrentBatch()`: Serializes and sends the current batch, updates tracking file, resets batch state.

## Execution Variables
- `SourcePath`: Root path for scanning the ITGlue export directory (Prompted).
- `Endpoint`: Target REST API endpoint for payload dispatch (Prompted).
- `ServiceNowUser` / `ServiceNowPass`: Credentials for API basic authentication (Prompted).
- `OversizedTrackingPath`: Path to `.migration-oversized.txt` in `$SourcePath`. Accumulates folder paths of all skipped oversized documents across runs.

## Developer / AI Notes
- **Arrays inside JSON Serialization**: PowerShell `ConvertTo-Json` natively flattens single-item arrays into objects. The script forces array casting `@(Get-DocumentAttachments...)` to ensure the `attachments` property consistently serializes as an `[]` array even with 1 attachment, honoring ServiceNow schema requirements.
- **Tracking / Resume**: The script writes `.migration-progress.txt` in `$SourcePath` containing the full path of the last successfully processed DOC- folder. On restart, if the file exists, the user is prompted to resume (skipping already-processed folders) or start fresh. The file is deleted upon successful batch completion.
- **Batching**: Documents are batched into a single payload (max 5MB). The `document` field is always an array. Size estimation uses string lengths to avoid unnecessary serialization.
- **Oversized document handling**: Documents whose estimated size exceeds 5 MB are silently dropped from batching (sending them would guarantee a 500 error). They are logged to `<CompanyName>.md` (human-readable, per-company markdown table) and `.migration-oversized.txt` (one folder path per line, machine-readable). The oversized tracking file is never auto-deleted — it persists across runs and is intended to seed a future re-run once the size limit is raised.

---

# AI Context: ITGlueImportDocuments (ServiceNow Script Include)

## System Metadata
- **Target File**: `x_wemop_knvl_itg_0/sys_script_include/ITGlueImportDocuments.script.js`
- **Primary Language**: ServiceNow JavaScript (ES5 / Rhino engine)
- **Core Purpose**: Receives JSON payloads from the PowerShell script and creates Knowledge Bases, Knowledge Articles, and attachments in ServiceNow. Also exposes a company validation endpoint.
- **Scope**: Scoped app `x_wemop_knvl_itg_0`

## Architecture & Execution Flow

### Instantiation
```js
initialize(mode)
```
- Sets `this.debug = (mode === "debug")`.
- All logging is gated on `this.debug`. When falsy, `_log` returns immediately — no output is written.

### `importDocuments(request)`
1. Parses `request.body.data` and validates the `document` array.
2. Iterates each document entry (`companyname`, `documentname`, `documentcontent`, `attachments[]`).
3. Calls `_getOrCreateKnowledgeBase(companyName)` → receives `{ kbSysId, domain }`.
4. Creates a `kb_knowledge` record in `draft` state, setting `sys_domain` if domain was resolved.
5. Re-fetches the article via `GlideRecord.get()` after insert, then sets `sys_domain` and updates.
6. Iterates attachments: calls `GlideSysAttachment.writeBase64()` per attachment.
7. Sets `workflow_state` to `published` and calls `update()`.
8. Collects per-document results `{ companyname, documentname, status, kb_sys_id, article_sys_id, attachments[] }`.

### `_getOrCreateKnowledgeBase(companyName)` → `{ kbSysId, domain }`
- Queries `kb_knowledge_base` by `title = companyName`.
- **Found**: returns `{ kbSysId: gr.getUniqueValue(), domain: gr.getValue("sys_domain") }`.
- **Not found**:
  1. Calls `checkCompany(null, companyName)` to resolve `sys_domain` from `core_company`.
  2. Inserts a new `kb_knowledge_base` record.
  3. Re-fetches the record, sets `sys_domain` if resolved, and calls `update()`.
  4. Returns `{ kbSysId: newKbSysId, domain: domain }`.

### `checkCompany(request, companyName)` — Dual-mode
- **Mode 1** — `checkCompany(null, companyName)`: Single internal lookup against `core_company`. Returns `gr.getValue("sys_domain")` or `null`. Used by `_getOrCreateKnowledgeBase`.
- **Mode 2** — `checkCompany(request)`: Bulk REST endpoint mode. Reads `request.body.data.companyNames[]`, queries each against `core_company`, returns `{ companyNames: [bool, ...] }`.

## Key Design Decisions
- **Domain propagation**: `sys_domain` flows from `core_company` → `kb_knowledge_base` → `kb_knowledge`. All three records end up in the same domain.
- **Return shape of `_getOrCreateKnowledgeBase`**: Returns an object `{ kbSysId, domain }` (not a plain string) so the caller has domain info without a second query.
- **`checkCompany` dual-mode**: The same function serves both the external REST API (bulk bool check) and internal domain resolution (single lookup), differentiated by whether `request` is null.
- **Debug mode**: All `_log` calls are no-ops unless instantiated with `"debug"`. Production callers should omit the argument to avoid log noise.
- **Article domain set post-insert**: `sys_domain` is applied via a re-fetch + `update()` after insert (not at `initialize()` time) to match the pattern established by the linter-adjusted code.
