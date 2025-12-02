BBKM Invoice Program - scripts only

## Authentication configuration

The scripts expect Microsoft Graph application credentials to be provided via
environment variables:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`

If you prefer to store these values in a local file rather than your shell
environment, set `AZURE_ENV_FILE` to point to a `.env`-style file containing
`KEY=VALUE` lines. When provided, the file's contents override any existing
environment values so fresh secrets (for example, a newly generated client
secret) are picked up immediately.

If `AZURE_ENV_FILE` is *not* set, the scripts also look for a `.env` file in
the same directory as the Python files themselves (for example, alongside
`Scripts/save_attachments_from_outlook_folder.py`). A sibling `.env` likewise
overrides existing values so deployments that rely on dropping a `.env` next to
the scripts always use the credentials in that file.
