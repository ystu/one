# Codex Notes

- Before starting any task in this workspace, explicitly read this `AGENTS.md` file and follow its instructions.
- This workspace contains Chinese Markdown files. When reading them in PowerShell, use UTF-8 explicitly to avoid mojibake:
  `[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new(); Get-Content -Raw -Encoding UTF8 -LiteralPath 'path'`
- When creating or editing speech manuscripts, refer to `docs/模板與提示詞/演講風格.md` and keep the writing aligned with that speaking style.
- When moving into each major heading or subheading in a speech manuscript, first add a brief spoken transition that bridges from the preceding topic into the next section, so the new heading is prepared by the prior content rather than appearing abruptly.
- When creating or editing outlines, follow `docs/模板與提示詞/大綱整理規則.md`.
- For any new speech topic, start by creating or refining the outline first. Confirm the content structure in the outline before expanding it into the speech manuscript.
- Speech manuscripts are expanded versions of their corresponding outlines. Treat each outline/manuscript pair as linked documents: when updating one, immediately review the other and apply any needed synchronized changes.
- For linked outline/manuscript pairs, the speech manuscript must include the outline's major headings and subheadings so the two documents can be read in parallel during preparation and delivery.
- For linked outline/manuscript pairs, keep both the major heading structure and subheading structure synchronized: the number, order, and meaning of corresponding headings should match unless there is an explicit reason to diverge. If the structures intentionally differ, state the reason clearly in the response.
- When revising either an outline or its paired speech manuscript, review whether the corresponding headings in the paired file also need to be updated, and apply synchronized changes immediately when needed.
- After editing any outline or speech manuscript, explicitly inspect the paired file before finishing. Verify the major heading count, subheading count, order, and meaning in both files. Do not assume no change is needed from memory; either synchronize the paired file or state the concrete reason for an intentional difference in the response.
