"""Convert a legacy .doc to .docx with a local Microsoft Word (Windows only).

A hidden Word instance with macros disabled opens a private copy read-only, without
dialogs, and saves it as DOCX. The caller passes the copy, never the original file.
"""

from __future__ import annotations

import os
import re
import subprocess
from datetime import UTC, datetime
from pathlib import Path

_VERSION_LINE = re.compile(r"^word_version=(\d{1,3}(?:\.\d{1,5}){0,3})\s*$", re.MULTILINE)
_CREATE_NO_WINDOW = 0x08000000

# Exit codes: 0 converted, 10 Word unavailable, 11 open failed, 12 save failed,
# 13 Word did not quit.
_SCRIPT = r"""
param([Parameter(Mandatory=$true)][string]$InputPath,
      [Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference = 'Stop'
try { $word = New-Object -ComObject Word.Application } catch { exit 10 }
$code = 0
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    try {
        $doc = $word.Documents.Open($InputPath, $false, $true, $false, 'no-password')
    } catch { $code = 11 }
    if ($code -eq 0) {
        try { $doc.SaveAs2($OutputPath, 16) } catch { $code = 12 }
        try { $doc.Close([ref]0) } catch {}
    }
    Write-Output ('word_version=' + $word.Version)
} finally {
    # Word's Quit and Close take by-reference arguments; a plain 0 throws in
    # PowerShell and would leave the hidden instance running.
    try { $word.Quit([ref]0) } catch { $code = 13 }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
exit $code
"""

_CLEANUP = r"""
param([Parameter(Mandatory=$true)][string]$Since)
$start = [datetime]::Parse($Since).ToUniversalTime()
Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
    Where-Object { $_.CommandLine -match '/Automation' -and
                   $_.CreationDate.ToUniversalTime() -ge $start } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
"""

_EXIT_REASONS = {10: "word_unavailable", 11: "word_open_failed", 12: "word_save_failed",
                 13: "word_quit_failed"}


class ConversionError(Exception):
    """Conversion failed; ``reason`` is a content-free code."""

    def __init__(self, reason: str) -> None:
        super().__init__(reason)
        self.reason = reason


def _powershell(script: Path, *args: str, timeout: int) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
         "-File", str(script), *args],
        capture_output=True, text=True, timeout=timeout, check=False,
        creationflags=_CREATE_NO_WINDOW if os.name == "nt" else 0,
    )


def convert_with_word(source: Path, target: Path, work_dir: Path, timeout: int) -> str | None:
    """Convert ``source`` (already a private copy) to DOCX at ``target``; return Word's version."""
    script = work_dir / "word_convert.ps1"
    cleanup = work_dir / "word_cleanup.ps1"
    script.write_text(_SCRIPT, encoding="utf-8-sig")
    cleanup.write_text(_CLEANUP, encoding="utf-8-sig")
    started = datetime.now(UTC).isoformat()
    try:
        # Word resolves relative paths against its own folder, not this process's.
        completed = _powershell(script, "-InputPath", str(source.resolve()),
                                "-OutputPath", str(target.resolve()), timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        _powershell(cleanup, "-Since", started, timeout=60)
        raise ConversionError("word_timeout") from exc
    except OSError as exc:
        raise ConversionError("powershell_unavailable") from exc
    if completed.returncode != 0:
        raise ConversionError(_EXIT_REASONS.get(completed.returncode, "word_failed"))
    if not target.is_file():
        raise ConversionError("word_output_missing")
    match = _VERSION_LINE.search(completed.stdout or "")
    return match.group(1) if match else None
