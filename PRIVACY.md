# Privacy

MultiDoc Sync processes documents locally. It does not upload file names,
paths, document contents, account names, browser data, or hardware identifiers.

## Optional anonymous usage statistics

When a release is connected to the project-operated telemetry endpoint, the
app asks for consent on first launch. The default choice is **No**. If enabled,
the app sends only:

- a randomly generated installation ID;
- a randomly generated session ID;
- app version, operating system, and CPU architecture;
- session start, five-minute heartbeat, and session end events;
- elapsed session seconds.

The installation ID is not derived from hardware or an account. The server
stores a keyed hash instead of the raw ID and does not persist IP addresses.
Telemetry can be disabled at any time from the **Privacy** button.

The public client contains no secret ingestion key. Abuse protection belongs
on the server through HTTPS, request validation, rate limits, and an
administrator-only statistics endpoint.
