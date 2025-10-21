from __future__ import annotations
import os
from dataclasses import dataclass
from dotenv import load_dotenv

load_dotenv()  # carrega .env da raiz

@dataclass(frozen=True)
class Settings:
    egestor_token: str = os.getenv("EGESTOR_PERSONAL_TOKEN", "").strip()

    imap_server: str = os.getenv("IMAP_SERVER", "").strip()
    imap_port: int = int(os.getenv("IMAP_PORT", "993"))
    email_user: str = os.getenv("EMAIL_USER", "").strip()
    email_pass: str = os.getenv("EMAIL_PASS", "").strip()

    tz: str = os.getenv("TZ", "America/Sao_Paulo")

    # limites/controle
    aio_concurrency: int = int(os.getenv("AIO_CONCURRENCY", "5"))
    http_timeout: int = int(os.getenv("HTTP_TIMEOUT", "30"))

settings = Settings()
