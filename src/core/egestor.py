from __future__ import annotations
import os
import aiohttp
import asyncio
import json
from pathlib import Path
from typing import Any, Dict, List, Optional
import requests
from .config import settings

API_BASE = "https://api.egestor.com.br/api"

def obter_token() -> dict:
    if not settings.egestor_token:
        raise RuntimeError("EGESTOR_PERSONAL_TOKEN ausente do .env")
    r = requests.post(
        f"{API_BASE}/oauth/access_token",
        headers={"Content-Type": "application/json"},
        json={"grant_type": "personal", "personal_token": settings.egestor_token},
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    return {
        "Authorization": f"Bearer {data['access_token']}",
        "Content-Type": "application/json",
    }

def listar_codigos_nfe(headers: dict, dt_ini: str, dt_fim: str) -> List[int]:
    codigos: List[int] = []
    page = 1
    while True:
        params = {
            "dtIni": dt_ini,
            "dtFim": dt_fim,
            "fields": "codigo",
            "orderBy": "codigo,asc",
            "limit": 100,
            "page": page,
        }
        r = requests.get(f"{API_BASE}/v1/nfe", headers=headers, params=params, timeout=30)
        r.raise_for_status()
        data = r.json().get("data", [])
        if not data:
            break
        codigos.extend(int(n["codigo"]) for n in data)
        page += 1
    return codigos

async def _baixar_uma_nfe(session: aiohttp.ClientSession, codigo: int, xml_dir: Path) -> Optional[dict]:
    try:
        async with session.get(f"{API_BASE}/v1/nfe/{codigo}") as resp:
            if resp.status != 200:
                return None
            detalhes = await resp.json()
    except Exception:
        return None

    try:
        async with session.get(f"{API_BASE}/v1/nfe/{codigo}/xml") as resp:
            if resp.status == 200:
                xml_bytes = await resp.read()
                xml_dir.mkdir(parents=True, exist_ok=True)
                (xml_dir / f"{codigo}.xml").write_bytes(xml_bytes)
                detalhes["__xml_saved__"] = True
            else:
                detalhes["__xml_saved__"] = False
    except Exception:
        detalhes["__xml_saved__"] = False

    return detalhes

async def baixar_lote_nfes(headers: dict, codigos: List[int], xml_dir: Path, *, concurrency: int = None, timeout: int = None) -> List[dict]:
    sem = asyncio.Semaphore(concurrency or settings.aio_concurrency)
    timeout = aiohttp.ClientTimeout(total=(timeout or settings.http_timeout))
    conn = aiohttp.TCPConnector(limit=(concurrency or settings.aio_concurrency))

    async def _guarded(c: int):
        async with sem:
            return await _baixar_uma_nfe(session, c, xml_dir)

    async with aiohttp.ClientSession(headers=headers, connector=conn, timeout=timeout) as session:
        tasks = [_guarded(c) for c in codigos]
        out = await asyncio.gather(*tasks, return_exceptions=True)

    return [o for o in out if isinstance(o, dict)]
