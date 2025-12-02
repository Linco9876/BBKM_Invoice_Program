BBKM Invoice Program - scripts only

## Authentication configuration

The scripts expect Microsoft Graph application credentials to be provided via
environment variables:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`

If you prefer to store these values in a local file rather than your shell
environment, set `AZURE_ENV_FILE` to point to a `.env`-style file containing
`KEY=VALUE` lines. Variables already present in the environment are not
overwritten, allowing you to keep secrets out of the repository while still
providing them to the program.

If `AZURE_ENV_FILE` is *not* set, the scripts also look for a `.env` file in
the same directory as the Python files themselves (for example, alongside
`Scripts/save_attachments_from_outlook_folder.py`). This makes it easy to ship a
local `.env` next to the scripts without exporting any variables manually.
