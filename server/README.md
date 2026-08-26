# Self-hosted anonymous telemetry API

This small Flask/SQLite service receives opt-in anonymous usage events. It does
not store IP addresses or raw installation IDs.

Required server-only environment variables:

```text
MULTIDOCSYNC_TELEMETRY_SALT=<long random value>
MULTIDOCSYNC_ADMIN_TOKEN=<long random value>
MULTIDOCSYNC_DB=/var/lib/multidoc-sync/telemetry.sqlite3
```

Never commit the real values. After deployment, set the public HTTPS ingestion
URL in `telemetry_config.py`. The URL is public and is not a secret.

Summary example:

```bash
curl -H "Authorization: Bearer $MULTIDOCSYNC_ADMIN_TOKEN" \
  https://aiwords.top/multidoc-sync-api/v1/admin/summary
```

The Nginx example disables access logs for this endpoint and adds rate limiting.
