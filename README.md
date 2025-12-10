BBKM Invoice Program - scripts only

## Authentication configuration

The scripts expect Microsoft Graph application credentials to be provided via
environment variables:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`

If you want to hardcode credentials directly in the scripts (for example while
troubleshooting), set the values in `HARD_CODED_AZURE_CREDENTIALS` near the top
of `Scripts/save_attachments_from_outlook_folder.py`. Those values override the
environment and any `.env` files when present.

If you prefer to store these values in a local file rather than your shell
environment, set `AZURE_ENV_FILE` to point to a `.env`-style file containing
`KEY=VALUE` lines. When provided, the file's contents override any existing
environment values so fresh secrets (for example, a newly generated client
secret) are picked up immediately.

If `AZURE_ENV_FILE` is *not* set, the scripts also look for a `.env` file in
the same directory as the Python files themselves (for example, alongside
`Scripts/save_attachments_from_outlook_folder.py`). A sibling `.env` likewise
overrides existing values so deployments that rely on dropping a `.env` next to
the scripts always use the credentials in that file. When credentials are
loaded, the script logs which source was chosen (environment, `.env`, or
`HARD_CODED_AZURE_CREDENTIALS`) along with the tenant and client ID being used
so you can verify the expected app registration is active when diagnosing
403/401 errors.

## Duplicate detection

`Scripts/save_attachments_from_outlook_folder.py` persists a small
`invoice_hashes.json` manifest alongside the script itself (within the
repository) instead of the attachment destination directory. Each saved
attachment is recorded by its SHA-256 hash so that any message with an
identical attachment will be categorised as **Doubled up** on future runs—even
if the category was manually removed or the original file was renamed. This
prevents already-seen invoices from being reprocessed without relying on
existing filenames alone.
