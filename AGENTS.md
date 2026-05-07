# Codex Notes

- This workspace contains Chinese Markdown files. When reading them in PowerShell, use UTF-8 explicitly to avoid mojibake:
  `[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new(); Get-Content -Raw -Encoding UTF8 -LiteralPath 'path'`
- When creating or editing speech manuscripts, refer to `docs/模板與提示詞/演講風格.md` and keep the writing aligned with that speaking style.
- When creating or editing outlines, follow `docs/模板與提示詞/大綱整理規則.md`.
- Speech manuscripts are expanded versions of their corresponding outlines. Treat each outline/manuscript pair as linked documents: when updating one, immediately review the other and apply any needed synchronized changes.
- For linked outline/manuscript pairs, keep the main heading structure synchronized: the number and order of major headings should match unless there is an explicit reason to diverge. If the structures intentionally differ, state the reason clearly in the response.
