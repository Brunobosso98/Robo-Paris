"""Shared configuration helpers that source sensitive values from .env."""

import os
from pathlib import Path


def _load_env(path: Path = None):
    """Load key/value pairs from a .env file without overriding existing env vars."""
    env_path = path or (Path(__file__).parent / ".env")
    if not env_path.exists():
        return

    with env_path.open(encoding="utf-8") as env_file:
        for line in env_file:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = map(str.strip, line.split("=", 1))
            if not key:
                continue
            if value.startswith(("'", '"')) and value.endswith(("'", '"')):
                value = value[1:-1]
            os.environ.setdefault(key, value)


def _require_env(variable_name: str) -> str:
    """Return the environment variable or raise if it is missing."""
    value = os.getenv(variable_name)
    if not value:
        raise EnvironmentError(f"Ambiente: variável obrigatória '{variable_name}' não configurada.")
    return value


_load_env()

SSPARISI_USERNAME = _require_env("SSPARISI_USERNAME")
SSPARISI_PASSWORD = _require_env("SSPARISI_PASSWORD")
RELACIONAMENTOS_USERNAME = _require_env("RELACIONAMENTOS_USERNAME")
RELACIONAMENTOS_PASSWORD = _require_env("RELACIONAMENTOS_PASSWORD")
