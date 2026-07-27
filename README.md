# PPT-TTS

Add AWS Polly text-to-speech narration to PowerPoint slides. Designed for Storyline 360 import — extracts slide notes, synthesizes speech, and embeds audio automatically.

## Features

- **GUI and CLI** — run with no arguments for the tkinter GUI, or pass arguments for headless use
- **Multi-language** — built-in presets for Mandarin (Zhiyu), English (Matthew), and Korean (Seoyeon)
- **Auto-detection** — detects language from filename characters
- **Pronunciation corrections** — per-voice replacement rules with preview
- **Slide selection** — process specific slides with ranges like `1,3,5-8`
- **Storyline normalization** — auto-normalizes embedded audio via PowerPoint COM for reliable Storyline 360 import

## Prerequisites

- Python 3.10+
- AWS account with Polly access
- (Windows only) Microsoft PowerPoint installed — used for audio normalization

## Setup

```bash
pip install -r requirements.txt
```

Copy `config.example.json` to `config.json` and fill in your AWS credentials:

```json
{
    "aws": {
        "access_key_id": "YOUR_AWS_ACCESS_KEY_ID",
        "secret_access_key": "YOUR_AWS_SECRET_ACCESS_KEY",
        "region": "us-east-1"
    }
}
```

Or configure AWS credentials via environment variables / `~/.aws/credentials` — the app will use them automatically.

## Usage

### GUI

```bash
python generate_audio_ppt.py
```

### CLI

```bash
python generate_audio_ppt.py presentation.pptx
```

```bash
python generate_audio_ppt.py presentation.pptx -o output.pptx --preset english --slides 1,3,5-8
```

### Options

| Flag | Description |
|------|-------------|
| `-o, --output` | Output file path (default: `narrated_<input>`) |
| `--preset` | Voice preset: `mandarin`, `english`, `korean` |
| `--voice` | Custom Polly voice ID (default: Zhiyu) |
| `--engine` | `neural` (default) or `standard` |
| `--lang` | Language code (default: `cmn-CN`) |
| `--slides` | Slide numbers/ranges, e.g. `1,3,5-8` |
| `--dry-run` | Show extracted notes without calling AWS |
| `--gui` | Force open the GUI |
| `-v, --verbose` | Debug logging |

### Voice Presets

| Preset | Voice | Language |
|--------|-------|----------|
| `mandarin` | Zhiyu | cmn-CN |
| `english` | Matthew | en-US |
| `korean` | Seoyeon | ko-KR |

## Building

To build a standalone executable:

```bash
pip install pyinstaller
pyinstaller PPT-TTS.spec
```

## License

See repository for license details.
