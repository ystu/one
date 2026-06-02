# Codex Notes

- Before starting any task in this workspace, explicitly read this `AGENTS.md` file and follow its instructions.
- This workspace contains Chinese Markdown files. When reading them in PowerShell, use UTF-8 explicitly to avoid mojibake:
  `[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new(); Get-Content -Raw -Encoding UTF8 -LiteralPath 'path'`
- When organizing meeting transcripts or meeting notes, first refer to `meeting/專有名詞表.md` and use its 「錯誤辨識／正確用詞」 mappings to correct names, titles, and proper nouns.
- When organizing meeting records, include the opening and closing remarks before and after the formal meeting when they appear in the transcript; place them near the beginning and end of the meeting record respectively.
- When creating or revising posters, first create or revise the no-text background/base visual and show it for approval. Do not add final typography, event details, or Chinese text until the background direction has been approved by the user.
- When creating or editing speech manuscripts, refer to `docs/模板與提示詞/演講風格.md` and keep the writing aligned with that speaking style.
- When creating or editing speech manuscripts, use this opening gratitude wording before introducing the topic: `敬愛的領班、操持講師、以及各位前賢大家晚安!大家好，後學承蒙上天慈悲、祖師弘慈、師恩母德、何老前人開荒台灣的五大犧牲、薛前人的浩然正氣、楊前人的老實修辦、王前人的好好修道、始終如一以及現今李前人慈悲的領導，永續創建一個六合圓滿、紀律化的優質道場，這些都要感恩過去前輩者的犧牲奉獻，建立了道場的基礎跟道場文化，才有今天這麼好的修道環境，後學也感念地區各位點傳師的提拔以及各位前賢們苦口婆心的成全與鼓勵，今天後學才有這個機會可以在這裡學習，後學今天要來學習的題目是OOOO！`
- When creating or editing speech manuscripts, use this closing wording at the end: `後學今日的學習講述到此告一段落，因為後學才疏學淺，還不能夠將道理講述得非常清楚，講述的過程中如有過失懇求上天慈悲，也懇請領班、操持講師以及各位前賢給予後學指正，祝福大家法喜充滿、聖凡如意，謝謝！`
- When moving into each major heading or subheading in a speech manuscript, first add a brief spoken transition that bridges from the preceding topic into the next section, so the new heading is prepared by the prior content rather than appearing abruptly.
- When creating or editing outlines, follow `docs/模板與提示詞/大綱整理規則.md`.
- For any new speech topic, start by creating or refining the outline first. Confirm the content structure in the outline before expanding it into the speech manuscript.
- Speech manuscripts are expanded versions of their corresponding outlines. Treat each outline/manuscript pair as linked documents: when updating one, immediately review the other and apply any needed synchronized changes.
- For linked outline/manuscript pairs, the speech manuscript must include the outline's major headings and subheadings so the two documents can be read in parallel during preparation and delivery.
- For linked outline/manuscript pairs, keep both the major heading structure and subheading structure synchronized: the number, order, and meaning of corresponding headings should match unless there is an explicit reason to diverge. If the structures intentionally differ, state the reason clearly in the response.
- When revising either an outline or its paired speech manuscript, review whether the corresponding headings in the paired file also need to be updated, and apply synchronized changes immediately when needed.
- After editing any outline or speech manuscript, explicitly inspect the paired file before finishing. Verify the major heading count, subheading count, order, and meaning in both files. Do not assume no change is needed from memory; either synchronize the paired file or state the concrete reason for an intentional difference in the response.
