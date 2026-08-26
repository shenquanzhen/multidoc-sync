from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKIP_DIRS = {".git", ".venv", "venv", "build", "dist", "runtime", "__pycache__"}
BINARY_SUFFIXES = {".ico", ".icns", ".png", ".jpg", ".jpeg", ".gif", ".pdf", ".exe", ".zip"}
SUSPICIOUS_NAMES = {
    ".env",
    ".env.local",
    "credentials.json",
    "cookies.txt",
    "id_rsa",
    "id_ed25519",
    "secrets.yml",
}

PATTERNS = {
    "private key": re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
    "GitHub token": re.compile(r"\b(?:ghp|gho|ghu|ghs|ghr)_[A-Za-z0-9]{20,}\b"),
    "OpenAI key": re.compile(r"\bsk-[A-Za-z0-9_-]{20,}\b"),
    "AWS access key": re.compile(r"\bAKIA[0-9A-Z]{16}\b"),
    "generic assigned secret": re.compile(
        r"(?i)\b(?:api[_-]?key|secret|token|password|passwd|cookie)\s*[:=]\s*['\"][^'\"]{8,}['\"]"
    ),
    "local Windows user path": re.compile(r"(?i)\bC:\\Users\\[^\\\s]+"),
    "local macOS user path": re.compile(r"/Users/[^/\s]+"),
}


def candidate_files():
    for path in ROOT.rglob("*"):
        if not path.is_file():
            continue
        if any(part in SKIP_DIRS for part in path.relative_to(ROOT).parts):
            continue
        yield path


def main() -> int:
    findings: list[str] = []
    scanned = 0
    for path in candidate_files():
        relative = path.relative_to(ROOT)
        if path.name.lower() in SUSPICIOUS_NAMES or path.suffix.lower() in {".pem", ".p12", ".pfx", ".key"}:
            findings.append(f"suspicious filename: {relative}")
        if path.suffix.lower() in BINARY_SUFFIXES or relative.as_posix() == "scripts/security_audit.py":
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            continue
        scanned += 1
        for label, pattern in PATTERNS.items():
            for match in pattern.finditer(text):
                line = text.count("\n", 0, match.start()) + 1
                findings.append(f"{label}: {relative}:{line}")

    if findings:
        print("Security audit failed:")
        for finding in findings:
            print(f"- {finding}")
        return 1
    print(f"Security audit passed: {scanned} text files checked; no common secrets or local user paths found.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
