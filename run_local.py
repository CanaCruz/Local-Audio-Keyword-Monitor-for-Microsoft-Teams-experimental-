"""
Alertas por palavra-chave SEM Microsoft Graph / Entra admin.

Você exporta ou guarda a transcrição como ficheiro (.vtt ou .txt) e corre este script.
Só precisa de TEAMS_INCOMING_WEBHOOK_URL no .env (webhook do canal — costuma dar
para criar sem ser admin global, se a política do Teams permitir).

Exemplos:
  python run_local.py transcricao.vtt
  python run_local.py transcricao.vtt --titulo "Daily 06/04"
  python run_local.py pasta\\*.vtt   (no PowerShell use Get-ChildItem e pipe, ver LEIAME)
"""

from __future__ import annotations

import argparse
import hashlib
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from dotenv import load_dotenv

# Reutiliza análise e envio do run_alerts.py (não executa o ciclo Graph).
from run_alerts import (
    KEYWORDS_FILE,
    append_log,
    already_sent,
    db_connect,
    dedup_key,
    find_hits,
    load_keywords,
    mark_sent,
    parse_vtt_cues,
    post_teams_webhook,
)

ROOT = Path(__file__).resolve().parent


def file_content_id(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:32]


def process_file(
    path: Path,
    webhook_url: str,
    titulo: str,
    conn,
    keywords: list[str],
) -> int:
    data = path.read_bytes()
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        text = data.decode("utf-8", errors="replace")

    tid = f"local:{file_content_id(data)}"
    meeting_label = titulo or path.stem
    when_hint = datetime.now(timezone.utc).isoformat()

    # .txt: tratar como um único bloco; .vtt: cues
    ext = path.suffix.lower()
    if ext == ".vtt" or "WEBVTT" in text[:20].upper():
        cues = parse_vtt_cues(text)
    else:
        cues = [("", line.strip()) for line in text.splitlines() if line.strip()]

    sent = 0
    for idx, (_ts, cue_text) in enumerate(cues):
        if not cue_text:
            continue
        for kw in find_hits(cue_text, keywords):
            h = dedup_key(tid, idx, kw)
            if already_sent(conn, h):
                continue
            try:
                post_teams_webhook(
                    webhook_url,
                    meeting_label,
                    when_hint,
                    kw,
                    cue_text,
                )
            except Exception as e:
                print(f"Erro ao enviar webhook: {e}", file=sys.stderr)
                continue
            mark_sent(conn, h)
            append_log(meeting_label, when_hint, kw, cue_text, tid)
            sent += 1
    return sent


def main() -> None:
    load_dotenv(ROOT / ".env")
    webhook = os.environ.get("TEAMS_INCOMING_WEBHOOK_URL", "").strip()
    if not webhook:
        print(
            "Defina TEAMS_INCOMING_WEBHOOK_URL no .env (único segredo obrigatório neste modo).",
            file=sys.stderr,
        )
        sys.exit(1)

    p = argparse.ArgumentParser(description="Alertas a partir de ficheiro local (sem Graph)")
    p.add_argument("ficheiro", type=Path, help="Caminho para .vtt ou .txt")
    p.add_argument("--titulo", default="", help="Nome da reunião no cartão do Teams")
    args = p.parse_args()

    path = args.ficheiro.resolve()
    if not path.is_file():
        print(f"Ficheiro não encontrado: {path}", file=sys.stderr)
        sys.exit(1)

    keywords = load_keywords(KEYWORDS_FILE)
    conn = db_connect()
    n = process_file(path, webhook, args.titulo.strip(), conn, keywords)
    print(f"OK — {path.name}: {n} alerta(s). Keywords: {', '.join(keywords)}")


if __name__ == "__main__":
    main()
