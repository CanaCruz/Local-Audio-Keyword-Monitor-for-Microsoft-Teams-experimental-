"""
================================================================================
MODO EXPERIMENTAL — ÁUDIO PELO WINDOWS (Stereo Mix / loopback)
================================================================================

Antes de usar na FIAP ou em reuniões com outras pessoas:
  - Política de uso aceitável da instituição pode proibir gravação/captura.
  - LGPD: tratar voz de terceiros exige base legal e transparência.
  - O recognize_google() ENVIA trechos de áudio para os servidores do Google.

Isto NÃO substitui a API oficial do Teams. É frágil (qualidade do áudio, ruído,
microfone errado, limites da API gratuita do Google).

Requisitos:
  1) Notificação no Teams:
     - Webhook clássico: canal → Conectores → Incoming Webhook → TEAMS_INCOMING_WEBHOOK_URL
     - Ou Power Automate: gatilho "Quando um pedido HTTP é recebido" → publicar no canal;
       copie o URL do pedido HTTP para TEAMS_POWER_AUTOMATE_HTTP_URL (JSON: keyword, excerpt, when_iso).
     Não use DRY_RUN=1 se quiser enviar de verdade.
  2) Windows: ativar "Stereo Mix" / "Mixagem estéreo" (Som > Gravação), OU usar
     VB-Audio Virtual Cable (saída do Teams → cabo; entrada do Python = "Cable Output"),
     OU definir AUDIO_INPUT_DEVICE_INDEX (vários: 32,30,29 — tenta por ordem).
  3) pip install -r requirements-experimental.txt

Uso:
  python experimental_listen_loopback.py

Teste sem enviar ao Teams (só imprime no terminal):
  set DRY_RUN=1
  python experimental_listen_loopback.py

Cada frase reconhecida (padrão): linha [STT] "…" → alerta / cooldown / nenhuma keyword.
  LOG_TRANSCRIPTS=0 desliga; STT_LANGUAGE=pt-BR ou en-US; KEYWORD_COOLDOWN_SEC=0 (padrão, sem cooldown; ex. 90 para espaçar alertas);
  SILENCE_MSG_EVERY=30 (menos avisos quando o áudio está em pausa).

Listar índices de microfone (não precisa de webhook):
  set LIST_AUDIO_DEVICES=1
  python experimental_listen_loopback.py

Rode num terminal normal (sessão com áudio). Jobs em background do PowerShell
costumam falhar ao abrir o microfone.

Mixagem estéreo: use LISTEN_MODE=chunk (padrão) — o modo vad/listen() por energia
quase não funciona com loopback. O som tem de sair na saída predefinida do Windows.

Headset USB (ex. Razer): a Mixagem estéreo Realtek muitas vezes NÃO mostra barras — o som
não passa pelo chip Realtek. Use LOOPBACK_MODE=wasapi + pip install PyAudioWPatch para
capturar direto a saída WASAPI (ex. "BlackShark ... Chat [Loopback]").

Se a Mixagem estéreo der erro PyAudio -9999: o script tenta modo sounddevice
(USE_WASAPI_LOOPBACK=1). O sounddevice 0.5.x não expõe loopback WASAPI —
nesse caso o script pede para usar AUDIO_INPUT_DEVICE_INDEX na Mixagem estéreo.
================================================================================
"""

from __future__ import annotations

import inspect
import os
import unicodedata
from typing import Any, Optional
import sys
import time
from datetime import datetime, timezone                                                                                             
from pathlib import Path

from dotenv import load_dotenv

try:
    import speech_recognition as sr
except ImportError:
    print("Instale: pip install SpeechRecognition", file=sys.stderr)
    sys.exit(1)

ROOT = Path(__file__).resolve().parent
load_dotenv(ROOT / ".env")

# Import pesado só depois da mensagem inicial (evita “tela preta” se falhar aqui).
def _lazy_import_run_alerts():
    from run_alerts import KEYWORDS_FILE, load_keywords, post_teams_webhook

    return KEYWORDS_FILE, load_keywords, post_teams_webhook

# Evita spam no canal (segundos entre alertas da mesma palavra-chave)
def _cooldown_sec() -> int:
    """0 ou negativo = sem cooldown (alerta em todo bloco em que a keyword aparecer)."""
    try:
        v = int(os.environ.get("KEYWORD_COOLDOWN_SEC", "0"))
    except ValueError:
        return 0
    return v


COOLDOWN_SEC = _cooldown_sec()

# A cada quantos segundos de “silêncio” mostra que o script ainda está vivo (0 = desliga)
HEARTBEAT_SEC = int(os.environ.get("HEARTBEAT_SEC", "30"))

# 1 = imprime no terminal tudo o que o Google transcrever (útil para ver se está ouvindo)
SHOW_HEARD = os.environ.get("SHOW_HEARD", "").strip().lower() in ("1", "true", "yes", "sim")

# 1 (padrão) = após cada bloco reconhecido: [STT] "texto…" → match: keyword ou nenhuma keyword
LOG_TRANSCRIPTS = os.environ.get("LOG_TRANSCRIPTS", "1").strip().lower() in (
    "1",
    "true",
    "yes",
    "sim",
)

# A cada quantos blocos silenciosos seguidos mostra aviso de RMS baixo (12 = frequente; suba se pausar música muito)
SILENCE_MSG_EVERY = max(1, int(os.environ.get("SILENCE_MSG_EVERY", "30")))

# Idioma do recognize_google (ex.: pt-BR, en-US) — se o Teams estiver noutro idioma, mude aqui ou no .env
def _stt_language() -> str:
    return (os.environ.get("STT_LANGUAGE", "pt-BR") or "pt-BR").strip()


# 1 = mostra timeouts e “Google não entendeu” (para perceber se o áudio chega)
DEBUG_STT = os.environ.get("DEBUG_STT", "").strip().lower() in ("1", "true", "yes", "sim")

# Multiplicador do limiar de energia após calibração (<1 = mais sensível). Ex.: 0.45
def _sensitivity_mult() -> float:
    try:
        v = float(os.environ.get("ENERGY_SENSITIVITY", "0.5"))
        return min(max(v, 0.1), 1.0)
    except ValueError:
        return 0.5


def list_input_devices() -> None:
    print("A ler a lista de microfones (demora um pouco no Windows)…", flush=True)
    print(
        "Dispositivos de entrada (use o índice em AUDIO_INPUT_DEVICE_INDEX no .env):",
        flush=True,
    )
    for i, name in enumerate(sr.Microphone.list_microphone_names()):
        print(f"  [{i}] {name!r}", flush=True)
    try:
        import pyaudiowpatch as pa_w

        pa = pa_w.PyAudio()
        try:
            print(
                "\n--- WASAPI [Loopback] (PyAudioWPatch — headset USB; LOOPBACK_MODE=wasapi) ---",
                flush=True,
            )
            for i in range(pa.get_device_count()):
                d = pa.get_device_info_by_index(i)
                if int(d.get("maxInputChannels") or 0) < 1:
                    continue
                if "loopback" not in (d.get("name") or "").lower():
                    continue
                print(f"  [{i}] {d.get('name', '')!r}", flush=True)
            try:
                d0 = pa.get_default_wasapi_loopback()
                print(
                    f"  Loopback da saída predefinida: índice {d0['index']} — {d0.get('name', '')!r}",
                    flush=True,
                )
            except Exception as ex:
                print(f"  (sem default WASAPI loopback: {ex})", flush=True)
        finally:
            pa.terminate()
    except ImportError:
        print(
            "\n(Dica headset USB: pip install PyAudioWPatch e LOOPBACK_MODE=wasapi no .env)\n",
            flush=True,
        )
    print(
        "\n--- Lista terminou. Para ESCUTAR, execute de novo SEM LIST_AUDIO_DEVICES, por exemplo:\n"
        "  Remove-Item Env:LIST_AUDIO_DEVICES -ErrorAction SilentlyContinue\n"
        "  $env:DRY_RUN = \"1\"; python -u experimental_listen_loopback.py\n",
        flush=True,
    )


def _bad_auto_input_name(low: str) -> bool:
    """Entradas que aparecem na lista mas não servem para ouvir o PC / reunião."""
    if "sound mapper" in low and "output" in low:
        return True
    if "mapper" in low and "output" in low and "input" not in low:
        return True
    return False


def _loopback_mode_wasapi_wpatch() -> bool:
    m = os.environ.get("LOOPBACK_MODE", "").strip().lower()
    return m in ("wasapi", "wpatch", "pyaudiowpatch")


def resolve_pyaudiowpatch_loopback() -> tuple[int, int, int, str]:
    """Índice PyAudioWPatch, taxa (Hz), canais e nome do dispositivo [Loopback]."""
    import pyaudiowpatch as pa_mod

    pa_mgr = pa_mod.PyAudio()
    try:
        idx_s = os.environ.get("WASAPI_LOOPBACK_DEVICE_INDEX", "").strip()
        if idx_s:
            i = int(idx_s)
            info = dict(pa_mgr.get_device_info_by_index(i))
            if int(info.get("maxInputChannels") or 0) < 1:
                raise OSError(f"O índice {i} não tem canais de entrada.")
            if "loopback" not in (info.get("name") or "").lower():
                print(
                    f"Aviso: o índice {i} não contém 'Loopback' no nome — pode ser microfone, não o som do PC.",
                    file=sys.stderr,
                    flush=True,
                )
        else:
            info = dict(pa_mgr.get_default_wasapi_loopback())
            i = int(info["index"])
        rate = int(float(info.get("defaultSampleRate") or 48000))
        ch = min(max(int(info.get("maxInputChannels") or 2), 1), 2)
        name = str(info.get("name", ""))
        return i, rate, ch, name
    finally:
        pa_mgr.terminate()


def guess_loopback_device_index() -> Optional[int]:
    print("A procurar Mixagem estéreo / loopback…", flush=True)
    names = sr.Microphone.list_microphone_names()
    # Ordem: nomes mais fiáveis primeiro (evita índice 3 “Sound Mapper Output”, etc.).
    primary = (
        "mixagem estéreo",
        "stereo mix",
        "wave out mix",
        "what u hear",
        "loopback",
        "cable output",
        "vb-audio",
        "voicemeeter",
    )
    secondary = ("mixagem", "stereo input")

    def scan(hint_tuple: tuple[str, ...]) -> Optional[int]:
        for i, name in enumerate(names):
            if not name:
                continue
            low = name.lower()
            if _bad_auto_input_name(low):
                continue
            if any(h in low for h in hint_tuple):
                return i
        return None

    idx = scan(primary)
    if idx is not None:
        return idx
    return scan(secondary)


def tune_recognizer_sensitivity(rec: sr.Recognizer) -> None:
    """Depois da calibração, o limiar costuma ficar alto — baixamos para apanhar fala mais baixa."""
    mult = _sensitivity_mult()
    before = rec.energy_threshold
    rec.energy_threshold = max(rec.energy_threshold * mult, 30.0)
    rec.pause_threshold = min(getattr(rec, "pause_threshold", 0.8), 0.6)
    rec.dynamic_energy_threshold = True
    print(
        f"Sensibilidade: limiar de energia {before:.0f} → {rec.energy_threshold:.0f} "
        f"(ENERGY_SENSITIVITY={mult}; ajuste 0.3–0.7 se precisar).",
        flush=True,
    )


def _parse_device_index_list(raw: str) -> list[int]:
    out: list[int] = []
    for part in raw.replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        try:
            out.append(int(part))
        except ValueError:
            continue
    return out


def resolve_device_index_candidates() -> list[int]:
    raw = os.environ.get("AUDIO_INPUT_DEVICE_INDEX", "").strip()
    if raw:
        cands = _parse_device_index_list(raw)
        if not cands:
            print(
                "AUDIO_INPUT_DEVICE_INDEX inválido. Use um número ou vários separados por vírgula: 32,30,29",
                file=sys.stderr,
                flush=True,
            )
            sys.exit(1)
        print(
            f"Índices a tentar (PyAudio): {cands} — usa o primeiro que abrir.",
            flush=True,
        )
        return cands
    idx = guess_loopback_device_index()
    if idx is not None:
        print(f"Usando dispositivo detectado automaticamente: índice {idx}", flush=True)
        return [idx]
    print(
        "Não achei Stereo Mix / loopback. Liste os índices abaixo, "
        "ative Mixagem estéreo no Painel de som e defina AUDIO_INPUT_DEVICE_INDEX no .env.\n",
        file=sys.stderr,
    )
    list_input_devices()
    sys.exit(1)


def heartbeat_line(device_index: int, hint: str = "") -> str:
    t = datetime.now().strftime("%H:%M:%S")
    if hint.strip():
        return f"[{t}] OK — script rodando ({hint.strip()})."
    return f"[{t}] OK — script rodando, ouvindo o dispositivo {device_index}."


def probe_working_input_params(device_index: int) -> tuple[int, int]:
    """
    Descobre taxa (Hz) e canais com que o PyAudio consegue abrir o dispositivo.
    Evita stream=None / erro no close com Mixagem estéreo (Realtek).
    """
    import pyaudio

    pa = pyaudio.PyAudio()
    try:
        info = pa.get_device_info_by_index(device_index)
        max_ch = int(info.get("maxInputChannels") or 0)
        if max_ch < 1:
            raise OSError(f"Dispositivo {device_index} não tem canais de entrada.")
        default_sr = int(float(info.get("defaultSampleRate") or 44100))
        rates: list[int] = []
        for r in (default_sr, 48000, 44100, 16000, 32000, 8000):
            if r not in rates:
                rates.append(r)
        ch_try = [2, 1] if max_ch >= 2 else [1]
        last_err: Exception | None = None
        for ch in ch_try:
            for rate in rates:
                stream = None
                try:
                    stream = pa.open(
                        format=pyaudio.paInt16,
                        channels=ch,
                        rate=rate,
                        input=True,
                        input_device_index=device_index,
                        frames_per_buffer=4096,
                    )
                    stream.stop_stream()
                    stream.close()
                    stream = None
                    print(
                        f"Áudio OK: dispositivo {device_index} a {rate} Hz, {ch} canal(is).",
                        flush=True,
                    )
                    return rate, ch
                except Exception as e:
                    last_err = e
                    if stream is not None:
                        try:
                            stream.close()
                        except Exception:
                            pass
        raise OSError(
            f"PyAudio não abriu o dispositivo {device_index}. "
            f"Último erro: {last_err}. Ative a Mixagem estéreo ou tente outro índice."
        ) from last_err
    finally:
        pa.terminate()


def build_microphone(device_index: int) -> sr.Microphone:
    sample_rate, _channels = probe_working_input_params(device_index)
    return sr.Microphone(
        device_index=device_index,
        sample_rate=sample_rate,
        chunk_size=4096,
    )


def run_pyaudio_chunk_loop(
    device_index: int,
    samplerate: int,
    channels: int,
    keywords: list[str],
    dry: bool,
    webhook: str,
    post_teams_webhook,
    rec: sr.Recognizer,
    *,
    pyaudio_pkg: Any = None,
    capture_backend: str = "stereo_mix",
    heartbeat_hint: str = "",
) -> None:
    """
    Lê blocos fixos de PCM (Mixagem estéreo ou WASAPI [Loopback] via PyAudioWPatch).
    O modo listen() por energia falha muito com loopback.
    """
    pyaudio = pyaudio_pkg or __import__("pyaudio")

    try:
        import numpy as np
    except ImportError:
        print("Modo chunk precisa de numpy: pip install numpy", file=sys.stderr, flush=True)
        sys.exit(1)

    chunk_sec = float(os.environ.get("PYAUDIO_CHUNK_SEC", "5"))
    frames = int(chunk_sec * samplerate)
    if frames < 2048:
        frames = 2048
    silent_rms = float(os.environ.get("SILENT_CHUNK_RMS", "35"))

    pa = pyaudio.PyAudio()
    stream = None
    try:
        stream = pa.open(
            format=pyaudio.paInt16,
            channels=channels,
            rate=samplerate,
            input=True,
            input_device_index=device_index,
            frames_per_buffer=frames,
        )
    except Exception as e:
        pa.terminate()
        print(f"Não abriu stream contínuo PyAudio: {e}", file=sys.stderr, flush=True)
        sys.exit(1)

    if capture_backend == "wasapi":
        print(
            f"Modo CHUNK + WASAPI loopback (PyAudioWPatch): ~{chunk_sec}s @ {samplerate} Hz, "
            f"{channels} canal(is). Ignora blocos com RMS < {silent_rms:.0f} (SILENT_CHUNK_RMS).",
            flush=True,
        )
    else:
        print(
            f"Modo CHUNK (Mixagem estéreo): ~{chunk_sec}s @ {samplerate} Hz, "
            f"{channels} canal(is). Ignora blocos com RMS < {silent_rms:.0f} (SILENT_CHUNK_RMS).",
            flush=True,
        )

    last_fire: dict[str, float] = {}
    last_heartbeat = time.monotonic()

    def ping_if_idle() -> None:
        nonlocal last_heartbeat
        if HEARTBEAT_SEC <= 0:
            return
        now = time.monotonic()
        if now - last_heartbeat >= HEARTBEAT_SEC:
            print(heartbeat_line(device_index, heartbeat_hint), flush=True)
            last_heartbeat = now

    unknown_streak = 0
    silent_streak = 0
    rms_zero_tip_shown = False

    try:
        while True:
            try:
                raw = stream.read(frames, exception_on_overflow=False)
            except Exception as ex:
                print(f"Erro ao ler áudio: {ex}", flush=True)
                time.sleep(0.5)
                ping_if_idle()
                continue

            arr = np.frombuffer(raw, dtype=np.int16)
            if channels > 1 and len(arr) >= channels:
                n = len(arr) - (len(arr) % channels)
                arr = arr[:n].reshape(-1, channels).mean(axis=1).astype(np.int16)
            rms = float(np.sqrt(np.mean(arr.astype(np.float64) ** 2)))
            if rms < silent_rms:
                silent_streak += 1
                if (
                    not rms_zero_tip_shown
                    and silent_streak >= 10
                    and rms < 1.0
                ):
                    rms_zero_tip_shown = True
                    if capture_backend == "wasapi":
                        print(
                            "\n*** RMS≈0 — WASAPI loopback sem sinal (nada a tocar nesta saída?) ***\n"
                            "Isto capta o que vai para os altifalantes/headset da entrada [Loopback] acima.\n\n"
                            "  1) Confirme que Teams/vídeo está a sair no MESMO dispositivo (Chat vs Game Razer).\n"
                            "  2) WASAPI_LOOPBACK_DEVICE_INDEX=31 força loopback do BlackShark **Game**; "
                            "omita para usar a saída predefinida do Windows.\n"
                            "  3) Volume de reprodução > 0; desative modo exclusivo nas propriedades do dispositivo.\n",
                            flush=True,
                        )
                    else:
                        print(
                            "\n*** RMS≈0 — a Mixagem estéreo NÃO está a receber o som do PC ***\n"
                            "A Mixagem grava o que sai na SAÍDA predefinida do Windows (não o microfone).\n"
                            "Com headset USB, a barra da Mixagem Realtek pode ficar parada — use então "
                            "LOOPBACK_MODE=wasapi + pip install PyAudioWPatch.\n\n"
                            "Checklist rápido:\n"
                            "  1) Reprodução (Painel de som): qual dispositivo tem o visto verde?\n"
                            "     O YouTube/Teams/tradutor tem de tocar AÍ (ex.: altifalantes Realtek ou headset).\n"
                            "  2) Gravação → Mixagem estéreo → Níveis: volume alto, NÃO mudos.\n"
                            "  3) Reprodução → o teu dispositivo → Avançado: desliga 'modo exclusivo'.\n"
                            "  4) Teste: abre um vídeo no browser COM SOM; o medidor da Mixagem deve mexer.\n"
                            "  5) Alguns drivers: só há sinal se houver som a sair mesmo (volume > 0).\n",
                            flush=True,
                        )
                if SHOW_HEARD and silent_streak % SILENCE_MSG_EVERY == 1:
                    if capture_backend == "wasapi":
                        print(
                            f"(SHOW_HEARD) Bloco silencioso (RMS={rms:.0f}). "
                            "Há som a tocar na saída capturada (Chat/Game)?",
                            flush=True,
                        )
                    else:
                        print(
                            f"(SHOW_HEARD) Bloco muito silencioso (RMS={rms:.0f}). "
                            "Suba o volume de saída ou confirme que o som vai para o dispositivo "
                            "que a Mixagem capta.",
                            flush=True,
                        )
                ping_if_idle()
                continue
            silent_streak = 0

            pcm = arr.astype(np.int16).tobytes()
            audio = sr.AudioData(pcm, samplerate, 2)

            try:
                texto = rec.recognize_google(audio, language=_stt_language())
            except sr.UnknownValueError:
                unknown_streak += 1
                if SHOW_HEARD or DEBUG_STT:
                    if unknown_streak % 2 == 1:
                        print(
                            f"(áudio com sinal RMS={rms:.0f}, mas Google não devolveu texto)",
                            flush=True,
                        )
                ping_if_idle()
                continue
            except sr.RequestError as e:
                print(f"API Google: {e}", file=sys.stderr, flush=True)
                time.sleep(3)
                ping_if_idle()
                continue

            unknown_streak = 0
            last_heartbeat = time.monotonic()
            fired, cooldowned = process_recognized_text(
                texto, keywords, dry, webhook, post_teams_webhook, last_fire
            )
            if LOG_TRANSCRIPTS or SHOW_HEARD:
                print(
                    f"[STT] {texto[:220]!r}  →  {_stt_result_line(fired, cooldowned)}",
                    flush=True,
                )
            ping_if_idle()
    finally:
        if stream is not None:
            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
        pa.terminate()


def _normalize_for_match(s: str) -> str:
    """Minúsculas + remove acentos para comparar com keywords."""
    if not s:
        return ""
    s = unicodedata.normalize("NFD", s.lower().strip())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def _keyword_in_normalized_text(text_norm: str, kw_norm: str) -> bool:
    if not kw_norm or not text_norm:
        return False
    if kw_norm in text_norm:
        return True
    # Nomes comuns mal transcritos
    if kw_norm in ("arthur", "artur"):
        return any(x in text_norm for x in ("arthur", "artur", "artu"))
    return False


def should_try_wasapi_fallback(exc: BaseException) -> bool:
    s = str(exc).lower()
    return (
        "9999" in s
        or "unanticipated" in s
        or "pyaudio não abriu" in s
        or ("nonetype" in s and "close" in s)
        or "stream de áudio é none" in s
    )


def _stt_result_line(fired: list[str], cooldowned: list[str]) -> str:
    """Resumo para a linha [STT] (alertas enviados vs. keyword no texto mas em cooldown)."""
    parts: list[str] = []
    if fired:
        parts.append("alerta: " + ", ".join(repr(k) for k in fired))
    if cooldowned:
        parts.append("cooldown: " + ", ".join(repr(k) for k in cooldowned))
    if parts:
        return " | ".join(parts)
    return "nenhuma keyword"


def post_power_automate_http(url: str, when_iso: str, keyword: str, excerpt: str) -> None:
    """POST JSON para o gatilho 'Quando um pedido HTTP é recebido' do Power Automate."""
    import requests

    payload = {
        "keyword": keyword,
        "excerpt": excerpt[:1800],
        "when_iso": when_iso,
    }
    r = requests.post(
        url,
        json=payload,
        headers={"Content-Type": "application/json; charset=utf-8"},
        timeout=60,
    )
    try:
        r.raise_for_status()
    except requests.HTTPError as e:
        if r.status_code == 401 and "sig=" not in url and "sig%3D" not in url:
            raise RuntimeError(
                "HTTP 401: URL do gatilho sem assinatura (sig). No passo HTTP, em «Quem pode disparar o fluxo?», "
                "experimente «Qualquer pessoa» (Anyone) — assim o portal costuma gerar URL com sig=/sp=/sv=. "
                "Guarde o fluxo, copie o URL COMPLETO nessa linha e atualize TEAMS_POWER_AUTOMATE_HTTP_URL no .env. "
                "Só ?api-version=1 não autentica o POST."
            ) from e
        raise


def process_recognized_text(
    texto: str,
    keywords: list[str],
    dry: bool,
    webhook: str,
    post_teams_webhook,
    last_fire: dict[str, float],
) -> tuple[list[str], list[str]]:
    """(keywords que dispararam alerta, keywords que batiam no texto mas estavam em cooldown)."""
    low = _normalize_for_match(texto)
    now = time.monotonic()
    when_iso = datetime.now(timezone.utc).isoformat()
    fired: list[str] = []
    cooldowned: list[str] = []
    for kw in keywords:
        kn = _normalize_for_match(kw)
        if not _keyword_in_normalized_text(low, kn):
            continue
        if COOLDOWN_SEC > 0:
            prev = last_fire.get(kw.lower(), 0)
            if now - prev < COOLDOWN_SEC:
                cooldowned.append(kw)
                continue
        last_fire[kw.lower()] = now
        fired.append(kw)
        banner = (
            f"\n{'=' * 58}\n"
            f"  ALERTA — palavra-chave: {kw!r}\n"
            f"  Trecho: {texto[:400]}{'…' if len(texto) > 400 else ''}\n"
            f"{'=' * 58}\n"
        )
        if dry:
            print(banner, flush=True)
            print(f"[DRY_RUN] Webhook não enviado (Teams). Texto acima.", flush=True)
            continue
        flow_url = os.environ.get("TEAMS_POWER_AUTOMATE_HTTP_URL", "").strip()
        try:
            if flow_url:
                post_power_automate_http(flow_url, when_iso, kw, texto)
                print(banner, flush=True)
                print(f"Alerta enviado (Power Automate HTTP): {kw!r}", flush=True)
            else:
                post_teams_webhook(
                    webhook,
                    f'Alerta áudio: palavra-chave "{kw}"',
                    when_iso,
                    kw,
                    texto,
                )
                print(banner, flush=True)
                print(f"Alerta enviado ao Teams (webhook): {kw!r}", flush=True)
        except Exception as ex:
            dest = "Power Automate" if flow_url else "webhook"
            print(f"Falha ao enviar ({dest}): {ex}", file=sys.stderr, flush=True)
    return fired, cooldowned


def _sounddevice_wasapi_loopback_device(sd: Any, default_out: int) -> Optional[Any]:
    """Só funciona se o sounddevice expuser WasapiSettings(loopback=...) (nem todas as versões têm)."""
    if not hasattr(sd, "WasapiSettings"):
        return None
    if "loopback" not in inspect.signature(sd.WasapiSettings).parameters:
        return None
    return (default_out, sd.WasapiSettings(loopback=True))


def _sounddevice_find_named_loopback_input(sd: Any) -> tuple[Optional[int], Optional[dict]]:
    """Alguns Windows/PortAudio listam entrada com 'loopback' no nome."""
    for i, d in enumerate(sd.query_devices()):
        try:
            ch = int(d.get("max_input_channels") or 0)
        except (TypeError, ValueError):
            continue
        if ch < 1:
            continue
        name = (d.get("name") or "").lower()
        if "loopback" in name:
            return i, d
    return None, None


def sounddevice_loopback_available() -> bool:
    """True se run_wasapi_loop conseguir abrir stream (API loopback ou entrada com 'loopback' no nome)."""
    try:
        import sounddevice as sd
    except ImportError:
        return False
    try:
        default_out = sd.default.device[1]
        if default_out is None or int(default_out) < 0:
            out_info = sd.query_devices(kind="output")
            default_out = int(out_info["index"])
        else:
            default_out = int(default_out)
        if _sounddevice_wasapi_loopback_device(sd, default_out) is not None:
            return True
        return _sounddevice_find_named_loopback_input(sd)[0] is not None
    except Exception:
        return False


def run_wasapi_loop(
    keywords: list[str],
    dry: bool,
    webhook: str,
    post_teams_webhook,
    rec: sr.Recognizer,
) -> None:
    import numpy as np

    try:
        import sounddevice as sd
    except ImportError:
        print(
            "Instale: pip install sounddevice numpy",
            file=sys.stderr,
            flush=True,
        )
        sys.exit(1)

    default_out = sd.default.device[1]
    if default_out is None or int(default_out) < 0:
        out_info = sd.query_devices(kind="output")
        default_out = int(out_info["index"])
    else:
        default_out = int(default_out)
        out_info = sd.query_devices(default_out, kind="output")

    chunk_sec = float(os.environ.get("WASAPI_CHUNK_SEC", "5"))

    stream_device: Any = None
    samplerate: int = 0
    channels: int = 1
    mode_label = "captura"

    loop_tuple = _sounddevice_wasapi_loopback_device(sd, default_out)
    if loop_tuple is not None:
        stream_device = loop_tuple
        samplerate = int(float(out_info["default_samplerate"]))
        channels = min(int(out_info["max_output_channels"]), 2)
        mode_label = "WASAPI loopback (API sounddevice)"
        print(
            f"{mode_label}: saída {out_info.get('name', '?')!r} @ {samplerate} Hz; "
            f"blocos ~{chunk_sec}s.",
            flush=True,
        )
    else:
        lb_i, lb_d = _sounddevice_find_named_loopback_input(sd)
        if lb_i is not None and lb_d is not None:
            stream_device = lb_i
            samplerate = int(float(lb_d.get("default_samplerate") or 48000))
            channels = min(int(lb_d.get("max_input_channels") or 1), 2)
            mode_label = f"entrada loopback [{lb_i}]"
            print(
                f"{mode_label}: {lb_d.get('name', '?')!r} @ {samplerate} Hz; "
                f"blocos ~{chunk_sec}s.",
                flush=True,
            )
        else:
            print(
                "USE_WASAPI_LOOPBACK não funciona com o seu sounddevice "
                f"{getattr(sd, '__version__', '?')}: não há WasapiSettings(loopback=) "
                "nem entrada com 'loopback' no nome.\n\n"
                "→ Use a Mixagem estéreo com PyAudio: no .env coloque\n"
                "  AUDIO_INPUT_DEVICE_INDEX=<índice>\n"
                "  (veja com LIST_AUDIO_DEVICES=1; ex. 30–32 em listas longas)\n"
                "e remova USE_WASAPI_LOOPBACK da sessão e do .env.\n\n"
                "Opcional (avançado): pip install PyAudioWPatch para loopback WASAPI real.",
                file=sys.stderr,
                flush=True,
            )
            sys.exit(1)

    frames = int(chunk_sec * samplerate)
    if frames < 1024:
        frames = 1024

    print(
        "(Som dos altifalantes predefinidos — deixa o Teams / TTS com áudio ligado.)",
        flush=True,
    )

    mode = "DRY_RUN (sem webhook)" if dry else "com webhook"
    print(f"Escutando… ({mode}) Ctrl+C para parar. Palavras:", ", ".join(keywords))
    if HEARTBEAT_SEC > 0:
        print(f"(Heartbeat a cada {HEARTBEAT_SEC}s)", flush=True)
    if SHOW_HEARD:
        print("(SHOW_HEARD=1 — cada frase reconhecida abaixo)", flush=True)

    last_fire: dict[str, float] = {}
    last_heartbeat = time.monotonic()

    def ping_if_idle() -> None:
        nonlocal last_heartbeat
        if HEARTBEAT_SEC <= 0:
            return
        now = time.monotonic()
        if now - last_heartbeat >= HEARTBEAT_SEC:
            t = datetime.now().strftime("%H:%M:%S")
            print(
                f"[{t}] OK — captura de sistema a correr ({mode_label}).",
                flush=True,
            )
            last_heartbeat = now

    try:
        with sd.InputStream(
            device=stream_device,
            channels=channels,
            samplerate=samplerate,
            dtype="float32",
        ) as stream:
            while True:
                try:
                    got = stream.read(frames)
                    if isinstance(got, tuple):
                        data = got[0]
                    else:
                        data = got
                except Exception as ex:
                    print(f"Erro ao ler áudio WASAPI: {ex}", flush=True)
                    time.sleep(1)
                    ping_if_idle()
                    continue

                mono = (
                    np.mean(data, axis=1)
                    if getattr(data, "ndim", 1) > 1 and data.shape[1] > 1
                    else np.asarray(data, dtype=np.float32).reshape(-1)
                )
                int16 = (np.clip(mono, -1.0, 1.0) * 32767.0).astype(np.int16)
                pcm = int16.tobytes()
                audio = sr.AudioData(pcm, samplerate, 2)

                try:
                    texto = rec.recognize_google(audio, language=_stt_language())
                except sr.UnknownValueError:
                    ping_if_idle()
                    continue
                except sr.RequestError as e:
                    print(f"API Google: {e}", file=sys.stderr, flush=True)
                    time.sleep(5)
                    ping_if_idle()
                    continue

                last_heartbeat = time.monotonic()
                fired, cooldowned = process_recognized_text(
                    texto, keywords, dry, webhook, post_teams_webhook, last_fire
                )
                if LOG_TRANSCRIPTS or SHOW_HEARD:
                    print(
                        f"[STT] {texto[:220]!r}  →  {_stt_result_line(fired, cooldowned)}",
                        flush=True,
                    )
                ping_if_idle()
    except OSError as e:
        print(
            f"WASAPI loopback falhou: {e}\n"
            "Use Windows 10/11, saída de áudio predefinida ativa, drivers atualizados.",
            file=sys.stderr,
            flush=True,
        )
        sys.exit(1)


def main() -> None:
    if os.environ.get("LIST_AUDIO_DEVICES"):
        list_input_devices()
        return

    KEYWORDS_FILE, load_keywords, post_teams_webhook = _lazy_import_run_alerts()

    dry = os.environ.get("DRY_RUN", "").strip() in ("1", "true", "yes", "sim")
    webhook = os.environ.get("TEAMS_INCOMING_WEBHOOK_URL", "").strip()
    flow_http = os.environ.get("TEAMS_POWER_AUTOMATE_HTTP_URL", "").strip()
    if not webhook and not flow_http and not dry:
        print(
            "Defina TEAMS_INCOMING_WEBHOOK_URL ou TEAMS_POWER_AUTOMATE_HTTP_URL no .env,\n"
            "ou use DRY_RUN=1 para testar só o áudio.",
            file=sys.stderr,
            flush=True,
        )
        sys.exit(1)

    print("A carregar keywords.txt…", flush=True)
    keywords = load_keywords(KEYWORDS_FILE)
    if not keywords:
        print("keywords.txt vazio.", file=sys.stderr)
        sys.exit(1)

    _notify = (
        "Power Automate HTTP"
        if flow_http
        else ("Teams webhook" if webhook else "—")
    )
    print(
        f"Config: STT={_stt_language()!r}, cooldown={'off' if COOLDOWN_SEC <= 0 else f'{COOLDOWN_SEC}s'}, "
        f"LOG_TRANSCRIPTS={'on' if LOG_TRANSCRIPTS else 'off'}, "
        f"DRY_RUN={'on' if dry else 'off'}, notificar={_notify}, SILENCE_MSG_EVERY={SILENCE_MSG_EVERY}.",
        flush=True,
    )

    rec = sr.Recognizer()
    rec.dynamic_energy_threshold = True

    if _loopback_mode_wasapi_wpatch():
        try:
            import pyaudiowpatch as pa_w
        except ImportError:
            print(
                "LOOPBACK_MODE=wasapi requer:\n  pip install PyAudioWPatch\n",
                file=sys.stderr,
                flush=True,
            )
            sys.exit(1)
        try:
            w_i, w_sr, w_ch, w_name = resolve_pyaudiowpatch_loopback()
        except Exception as e:
            print(
                f"Não foi possível abrir WASAPI loopback: {e}\n"
                "Liste índices [Loopback] com LIST_AUDIO_DEVICES=1 (precisa de PyAudioWPatch).",
                file=sys.stderr,
                flush=True,
            )
            sys.exit(1)
        hb = f"WASAPI loopback índice {w_i}: {w_name!r}"
        mode = "DRY_RUN (sem webhook)" if dry else "com webhook"
        print(
            f"LOOPBACK_MODE=wasapi — áudio da saída: {w_name!r} (índice {w_i}).",
            flush=True,
        )
        print(
            "Útil para headset USB: não depende da Mixagem estéreo Realtek.",
            flush=True,
        )
        print(f"Escutando… ({mode}) Ctrl+C para parar. Palavras:", ", ".join(keywords))
        if HEARTBEAT_SEC > 0:
            print(
                f"(Heartbeat a cada {HEARTBEAT_SEC}s — {heartbeat_line(w_i, hb)})",
                flush=True,
            )
        if SHOW_HEARD:
            print("(SHOW_HEARD=1 — texto reconhecido ou avisos de silêncio)", flush=True)
        run_pyaudio_chunk_loop(
            w_i,
            w_sr,
            w_ch,
            keywords,
            dry,
            webhook,
            post_teams_webhook,
            rec,
            pyaudio_pkg=pa_w,
            capture_backend="wasapi",
            heartbeat_hint=hb,
        )
        return

    force_wasapi = os.environ.get("USE_WASAPI_LOOPBACK", "").strip().lower() in (
        "1",
        "true",
        "yes",
        "sim",
    )
    if force_wasapi:
        if sounddevice_loopback_available():
            print("USE_WASAPI_LOOPBACK=1 — modo sounddevice (loopback disponível).", flush=True)
            run_wasapi_loop(keywords, dry, webhook, post_teams_webhook, rec)
            return
        print(
            "USE_WASAPI_LOOPBACK=1 está ligado, mas este sounddevice não tem loopback WASAPI útil.\n"
            "→ A continuar com PyAudio + Mixagem estéreo (AUDIO_INPUT_DEVICE_INDEX no .env).\n"
            "→ Para não voltar a aparecer isto: Remove-Item Env:USE_WASAPI_LOOPBACK",
            flush=True,
        )

    print("A escolher o dispositivo de áudio…", flush=True)
    candidates = resolve_device_index_candidates()

    print("A testar abertura com PyAudio (taxa de amostragem)…", flush=True)
    mic = None
    device_index = -1
    last_err: Optional[BaseException] = None
    for idx in candidates:
        try:
            print(f"  Índice {idx}…", flush=True)
            mic = build_microphone(idx)
            device_index = idx
            print(f"  → Dispositivo {idx} abriu com sucesso.", flush=True)
            break
        except Exception as e:
            last_err = e
            print(f"  → Falhou: {e}", flush=True)

    if mic is None:
        e = last_err or RuntimeError("sem candidatos")
        if should_try_wasapi_fallback(e):
            if sounddevice_loopback_available():
                print(
                    "PyAudio falhou em todos os índices. A tentar sounddevice (loopback)…",
                    flush=True,
                )
                run_wasapi_loop(keywords, dry, webhook, post_teams_webhook, rec)
                return
            print(
                "PyAudio não abriu nenhum índice e sounddevice não tem loopback neste PC.\n\n"
                "Faça:\n"
                "  1) Painel de som → Gravação → ativar 'Mixagem estéreo' e subir o volume.\n"
                "  2) No .env: AUDIO_INPUT_DEVICE_INDEX=32,30,29 (vários à prova).\n"
                "  3) PowerShell: Remove-Item Env:USE_WASAPI_LOOPBACK\n"
                "  4) Headset USB: pip install PyAudioWPatch e no .env LOOPBACK_MODE=wasapi\n\n"
                f"Último erro: {e}",
                file=sys.stderr,
                flush=True,
            )
            sys.exit(1)
        print(
            f"Falha ao preparar o microfone (índices tentados: {candidates}).\n"
            f"Erro: {e}",
            file=sys.stderr,
            flush=True,
        )
        sys.exit(1)

    listen_mode = os.environ.get("LISTEN_MODE", "chunk").strip().lower()
    if listen_mode not in ("vad", "energy", "listen"):
        sr_ch, ch_ch = probe_working_input_params(device_index)
        mode = "DRY_RUN (sem webhook)" if dry else "com webhook"
        print(
            "LISTEN_MODE=chunk — lê blocos de áudio (funciona melhor com Mixagem estéreo que o modo VAD).",
            flush=True,
        )
        print(
            "Dica: o som do PC/Teams/TTS tem de ir para a **saída de áudio predefinida** do Windows; "
            "a Mixagem capta essa saída. Suba o volume de reprodução.",
            flush=True,
        )
        print(f"Escutando… ({mode}) Ctrl+C para parar. Palavras:", ", ".join(keywords))
        if HEARTBEAT_SEC > 0:
            print(
                f"(Heartbeat a cada {HEARTBEAT_SEC}s — {heartbeat_line(device_index)})",
                flush=True,
            )
        if SHOW_HEARD:
            print("(SHOW_HEARD=1 — texto reconhecido ou avisos de silêncio)", flush=True)
        run_pyaudio_chunk_loop(
            device_index,
            sr_ch,
            ch_ch,
            keywords,
            dry,
            webhook,
            post_teams_webhook,
            rec,
        )
        return

    print("Calibrando ruído ambiente (1 s)…", flush=True)
    try:
        with mic as source:
            if source.stream is None:
                raise OSError("Stream de áudio é None — dispositivo não abriu.")
            rec.adjust_for_ambient_noise(source, duration=1.0)
            tune_recognizer_sensitivity(rec)
    except Exception as e:
        if should_try_wasapi_fallback(e) and sounddevice_loopback_available():
            print(
                "Calibração PyAudio falhou. A tentar sounddevice (loopback)…",
                flush=True,
            )
            run_wasapi_loop(keywords, dry, webhook, post_teams_webhook, rec)
            return
        if should_try_wasapi_fallback(e):
            print(
                "Calibração PyAudio falhou e sounddevice não tem loopback. "
                "Ajuste AUDIO_INPUT_DEVICE_INDEX.",
                file=sys.stderr,
                flush=True,
            )
            sys.exit(1)
        print(
            f"Falha na calibração (índice {device_index}). Erro: {e}",
            file=sys.stderr,
            flush=True,
        )
        sys.exit(1)

    last_fire: dict[str, float] = {}
    mode = "DRY_RUN (sem webhook)" if dry else "com webhook"
    print(
        "Dica: som dos altifalantes → Mixagem estéreo (índice no .env) ou loopback sounddevice se disponível.",
        flush=True,
    )
    print(f"Escutando… ({mode}) Ctrl+C para parar. Palavras:", ", ".join(keywords))
    if HEARTBEAT_SEC > 0:
        print(
            f"(A cada {HEARTBEAT_SEC}s sem atividade aparece: "
            f'"{heartbeat_line(device_index)}")',
            flush=True,
        )
    if SHOW_HEARD:
        print("(SHOW_HEARD=1 — cada frase reconhecida aparece abaixo)", flush=True)
    if DEBUG_STT:
        print(
            "(DEBUG_STT=1 — avisos de timeout / Google sem texto)",
            flush=True,
        )

    last_heartbeat = time.monotonic()
    dbg = {"timeout_streak": 0, "unknown_streak": 0}

    def ping_if_idle() -> None:
        nonlocal last_heartbeat
        if HEARTBEAT_SEC <= 0:
            return
        now = time.monotonic()
        if now - last_heartbeat >= HEARTBEAT_SEC:
            print(heartbeat_line(device_index), flush=True)
            last_heartbeat = now

    while True:
        try:
            with mic as source:
                audio = rec.listen(source, timeout=6, phrase_time_limit=18)
        except sr.WaitTimeoutError:
            dbg["timeout_streak"] += 1
            if DEBUG_STT and dbg["timeout_streak"] % 5 == 1:
                print(
                    "(DEBUG) Timeout — nenhuma fala forte o suficiente neste microfone. "
                    "Suba o volume ou mude de dispositivo / USE_WASAPI_LOOPBACK=1.",
                    flush=True,
                )
            ping_if_idle()
            continue
        except OSError as e:
            print(f"Erro de áudio: {e}", file=sys.stderr)
            time.sleep(2)
            ping_if_idle()
            continue

        try:
            texto = rec.recognize_google(audio, language=_stt_language())
        except sr.UnknownValueError:
            dbg["unknown_streak"] += 1
            dbg["timeout_streak"] = 0
            if DEBUG_STT and dbg["unknown_streak"] % 3 == 1:
                print(
                    "(DEBUG) Google não reconheceu texto neste bloco (ruído, volume ou idioma).",
                    flush=True,
                )
            ping_if_idle()
            continue
        except sr.RequestError as e:
            print(f"API Google: {e}", file=sys.stderr)
            time.sleep(5)
            ping_if_idle()
            continue

        last_heartbeat = time.monotonic()
        dbg["unknown_streak"] = 0
        dbg["timeout_streak"] = 0
        fired, cooldowned = process_recognized_text(
            texto, keywords, dry, webhook, post_teams_webhook, last_fire
        )
        if LOG_TRANSCRIPTS or SHOW_HEARD:
            print(
                f"[STT] {texto[:220]!r}  →  {_stt_result_line(fired, cooldowned)}",
                flush=True,
            )
        ping_if_idle()


if __name__ == "__main__":
    try:
        _reconf_kw: dict = {"line_buffering": True}
        if sys.platform == "win32":
            _reconf_kw.update(encoding="utf-8", errors="replace")
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(**_reconf_kw)
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(**_reconf_kw)
    except Exception:
        pass
    print(
        "experimental_listen_loopback.py — se ficar parado aqui, o áudio do Windows pode estar a bloquear.",
        flush=True,
    )
    try:
        main()
    except KeyboardInterrupt:
        print("\nEncerrado.", flush=True)
