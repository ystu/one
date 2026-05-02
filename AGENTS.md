# Codex Notes

- This workspace contains Chinese Markdown files. When reading them in PowerShell, use UTF-8 explicitly to avoid mojibake:
  `[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new(); Get-Content -Raw -Encoding UTF8 -LiteralPath 'path'`
- When creating or editing speech/manuscript files, refer to `docs/模板與提示詞/演講風格.md` and keep the writing aligned with that speaking style.
- Speech manuscripts and outlines are linked documents. When updating one, check the corresponding file and keep both synchronized.
