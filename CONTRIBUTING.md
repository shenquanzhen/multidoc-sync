# Contributing

Issues and pull requests are welcome. Keep document handling local and avoid
adding telemetry, cloud uploads, credential collection, or source-file writes.

Before submitting a pull request:

1. Run `python -m unittest discover -s tests -v`.
2. Test PDF mode with both three and four files.
3. If changing native Office control, test with read-only copies on the target OS.
4. Do not commit documents, build output, logs, cookies, tokens, or machine-specific paths.
