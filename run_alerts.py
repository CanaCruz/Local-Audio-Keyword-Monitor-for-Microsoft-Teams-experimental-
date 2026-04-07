"""
Sincroniza transcrições (Microsoft Graph delta), procura keywords e envia alertas ao Teams.

Requisitos no Entra ID (app registration):
  - Application permission: OnlineMeetingTranscript.Read.All
  - Admin consent

Para permissões de aplicação em reuniões online, a Microsoft exige também uma
**application access policy** a associar ao utilizador cujo ID usas em MEETING_ORGANIZER_USER_ID.
Ver: "Allow applications to access online meetings on behalf of a user"

Uso:
  python run_alerts.py              # uma execução (ex.: Agendador do Windows / cron)
  python run_alerts.py --loop 300   # repete a cada 300 segundos
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import os
import re
import sqlite3
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, Optional
from urllib.parse import quote

import requests
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent
STATE_DIR = ROOT / "state"
KEYWORDS_FILE = ROOT / "keywords.txt"
DELTA_STATE_FILE = STATE_DIR / "delta_link.txt"
DB_FILE = STATE_DIR / "dedup.sqlite"
LOG_CSV = STATE_DIR / "alerts_log.csv"

GRAPH = "https://graph.microsoft.com/v1.0"


def load_keywords(path: Path) -> list[str]:
    if not path.is_file():
        return ["chamada", "arthur"]
    out: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        out.append(line)
    return out if out else ["chamada", "arthur"]


def get_token(tenant: str, client_id: str, client_secret: str) -> str:
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data, timeout=60)
    r.raise_for_status()
    return r.json()["access_token"]


def initial_delta_url(user_id: str) -> str:
    # meetingOrganizerUserId é obrigatório (documentação Graph).
    oid = quote(user_id, safe="")
    return (
        f"{GRAPH}/users/{oid}/onlineMeetings/"
        f"getAllTranscripts(meetingOrganizerUserId='{user_id}')/delta"
    )


def read_delta_url() -> Optional[str]:
    if not DELTA_STATE_FILE.is_file():
        return None
    u = DELTA_STATE_FILE.read_text(encoding="utf-8").strip()
    return u or None


def write_delta_url(url: str) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    DELTA_STATE_FILE.write_text(url, encoding="utf-8")


def fetch_delta_round(token: str, start_url: str) -> tuple[list[dict], str]:
    """Segue nextLink até obter deltaLink; devolve (items, novo_delta_link)."""
    items: list[dict] = []
    url = start_url
    headers = {"Authorization": f"Bearer {token}"}
    final_delta: str | None = None

    while url:
        r = requests.get(url, headers=headers, timeout=120)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph error {r.status_code}: {r.text[:2000]}")
        data = r.json()
        for obj in data.get("value", []):
            if obj.get("@removed") is not None:
                continue
            items.append(obj)
        next_link = data.get("@odata.nextLink")
        delta_link = data.get("@odata.deltaLink")
        if next_link:
            url = next_link
            continue
        if delta_link:
            final_delta = delta_link
            url = None
        else:
            raise RuntimeError("Resposta delta sem nextLink nem deltaLink.")

    if not final_delta:
        raise RuntimeError("Sem deltaLink na última página.")
    return items, final_delta


def strip_vtt_text(line: str) -> str:
    line = re.sub(r"<[^>]+>", " ", line)
    return html.unescape(line)


def parse_vtt_cues(vtt: str) -> list[tuple[str, str]]:
    """Lista de (timestamp_line, texto)."""
    cues: list[tuple[str, str]] = []
    blocks = re.split(r"\n\n+", vtt.strip())
    for block in blocks:
        lines = [ln.rstrip() for ln in block.splitlines() if ln.strip()]
        if not lines:
            continue
        if lines[0].strip().isdigit():
            lines = lines[1:]
        if not lines:
            continue
        if "-->" in lines[0]:
            ts = lines[0].strip()
            text = " ".join(strip_vtt_text(x) for x in lines[1:])
        else:
            ts = ""
            text = " ".join(strip_vtt_text(x) for x in lines)
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            cues.append((ts, text))
    return cues


def find_hits(text: str, keywords: Iterable[str]) -> list[str]:
    t = text.lower()
    hits: list[str] = []
    for kw in keywords:
        if kw.lower() in t:
            hits.append(kw)
    return hits


def dedup_key(transcript_id: str, cue_index: int, keyword: str) -> str:
    raw = f"{transcript_id}|{cue_index}|{keyword.lower()}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def db_connect() -> sqlite3.Connection:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS seen (h TEXT PRIMARY KEY, ts TEXT NOT NULL)"
    )
    conn.commit()
    return conn


def already_sent(conn: sqlite3.Connection, h: str) -> bool:
    row = conn.execute("SELECT 1 FROM seen WHERE h = ?", (h,)).fetchone()
    return row is not None


def mark_sent(conn: sqlite3.Connection, h: str) -> None:
    now = datetime.now(timezone.utc).isoformat()
    conn.execute(
        "INSERT OR IGNORE INTO seen (h, ts) VALUES (?, ?)", (h, now)
    )
    conn.commit()


def append_log(
    meeting_label: str,
    when_iso: str,
    keyword: str,
    excerpt: str,
    transcript_id: str,
) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    new_file = not LOG_CSV.is_file()
    with LOG_CSV.open("a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(
                ["utc_time", "meeting", "datetime_hint", "keyword", "excerpt", "transcript_id"]
            )
        w.writerow(
            [
                datetime.now(timezone.utc).isoformat(),
                meeting_label,
                when_iso,
                keyword,
                excerpt[:4000],
                transcript_id,
            ]
        )


def get_meeting_subject(token: str, user_id: str, meeting_id: str) -> Optional[str]:
    mid = quote(meeting_id, safe="")
    uid = quote(user_id, safe="")
    url = f"{GRAPH}/users/{uid}/onlineMeetings/{mid}?$select=subject,startDateTime,endDateTime"
    r = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=60,
    )
    if r.status_code != 200:
        return None
    data = r.json()
    return data.get("subject") or None


def post_teams_webhook(
    webhook_url: str,
    meeting_name: str,
    when_iso: str,
    keyword: str,
    excerpt: str,
) -> None:
    card = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "summary": "Alerta: palavra-chave na transcrição",
        "themeColor": "0078D4",
        "title": meeting_name,
        "sections": [
            {
                "facts": [
                    {"name": "Data/hora (referência)", "value": when_iso},
                    {"name": "Palavra-chave", "value": keyword},
                ],
                "text": excerpt[:1800],
            }
        ],
    }
    r = requests.post(
        webhook_url,
        data=json.dumps(card).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=utf-8"},
        timeout=60,
    )
    r.raise_for_status()


def process_transcripts(
    token: str,
    organizer_id: str,
    items: list[dict],
    keywords: list[str],
    webhook_url: str,
    conn: sqlite3.Connection,
) -> int:
    sent = 0
    for tr in items:
        tid = tr.get("id")
        meeting_id = tr.get("meetingId")
        content_url = tr.get("transcriptContentUrl")
        created = tr.get("createdDateTime") or ""
        if not tid or not content_url or not meeting_id:
            continue

        r = requests.get(
            content_url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=120,
        )
        if r.status_code != 200:
            continue
        vtt = r.text
        subject = get_meeting_subject(token, organizer_id, meeting_id)
        meeting_label = subject or meeting_id

        cues = parse_vtt_cues(vtt)
        for idx, (ts_line, cue_text) in enumerate(cues):
            for kw in find_hits(cue_text, keywords):
                h = dedup_key(tid, idx, kw)
                if already_sent(conn, h):
                    continue
                when_hint = created or ts_line or ""
                try:
                    post_teams_webhook(
                        webhook_url,
                        meeting_label,
                        when_hint,
                        kw,
                        cue_text,
                    )
                except requests.RequestException:
                    continue
                mark_sent(conn, h)
                append_log(meeting_label, when_hint, kw, cue_text, tid)
                sent += 1
    return sent


def run_cycle() -> None:
    load_dotenv(ROOT / ".env")
    tenant = os.environ.get("AZURE_TENANT_ID", "").strip()
    cid = os.environ.get("AZURE_CLIENT_ID", "").strip()
    secret = os.environ.get("AZURE_CLIENT_SECRET", "").strip()
    org = os.environ.get("MEETING_ORGANIZER_USER_ID", "").strip()
    webhook = os.environ.get("TEAMS_INCOMING_WEBHOOK_URL", "").strip()

    for name, val in [
        ("AZURE_TENANT_ID", tenant),
        ("AZURE_CLIENT_ID", cid),
        ("AZURE_CLIENT_SECRET", secret),
        ("MEETING_ORGANIZER_USER_ID", org),
        ("TEAMS_INCOMING_WEBHOOK_URL", webhook),
    ]:
        if not val:
            print(f"Defina {name} no ficheiro .env (veja .env.example).", file=sys.stderr)
            sys.exit(1)

    keywords = load_keywords(KEYWORDS_FILE)
    token = get_token(tenant, cid, secret)

    start = read_delta_url() or initial_delta_url(org)
    items, new_delta = fetch_delta_round(token, start)
    write_delta_url(new_delta)

    conn = db_connect()
    n = process_transcripts(token, org, items, keywords, webhook, conn)
    print(
        f"OK — {len(items)} transcrição(ões) neste ciclo, {n} alerta(s) enviado(s). "
        f"Keywords: {', '.join(keywords)}"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Alertas Teams por keywords em transcrições")
    parser.add_argument(
        "--loop",
        type=int,
        metavar="SEG",
        default=0,
        help="Repetir a cada SEG segundos (0 = executar uma vez)",
    )
    args = parser.parse_args()

    if args.loop and args.loop < 30:
        print("--loop deve ser >= 30 para não sobrecarregar a API.", file=sys.stderr)
        sys.exit(1)

    while True:
        run_cycle()
        if not args.loop:
            break
        time.sleep(args.loop)


if __name__ == "__main__":
    main()
