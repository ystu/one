# Codex Notes

- Before starting any task in this workspace, explicitly read this `AGENTS.md` file and follow its instructions.
- Reading Chinese Markdown in PowerShell: set UTF-8 output and use `-Encoding UTF8`, e.g.
  `[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new(); Get-Content -Raw -Encoding UTF8 -LiteralPath 'path'`
- When organizing or editing transcripts, outlines, speech manuscripts, or other documents related to the 道場, refer to `docs/專有名詞表.md` and use its mappings and confirmed terms to correct names, titles, and 道場-specific terminology.
- When organizing transcripts, first create or revise only the Markdown file for the user to review. Do not generate, regenerate, or update a Word `.docx` file until the user has reviewed the Markdown and explicitly confirms that a Word version is needed.
- Determine transcript formatting by content type. For meeting transcripts, organize the content into clear, readable paragraphs. For an individual speech transcript, group the original text into natural paragraphs by complete thought, story, or topic transition so it remains easy to read, but do not add topical headings unless the user explicitly requests them.
- When organizing meeting transcripts or meeting notes, first refer to `meeting/專有名詞表.md` and use its 「錯誤辨識／正確用詞」 mappings to correct names, titles, and proper nouns.
- When organizing meeting records, include the opening and closing remarks before and after the formal meeting when they appear in the transcript; place them near the beginning and end of the meeting record respectively.
- When creating or revising posters, first create or revise the no-text background/base visual and show it for approval. Do not add final typography, event details, or Chinese text until the background direction has been approved by the user.
- When creating or editing speech manuscripts, refer to `docs/模板與提示詞/演講風格.md` and keep the writing aligned with that speaking style.
- 撰寫或修改講稿時，所有實質內容（包括道理說明、經典引文、人物言論、歷史事蹟、故事、案例及數據）都必須以目前 repo 內已實際查閱的資料為依據，不得憑模型記憶、網路資料或推測自行補寫。可在忠於原意的前提下整理、口語化及添加銜接語，但不得新增來源未支持的主張或細節。
- 講稿的重要主張、引文及故事，撰寫前須核對 repo 內已實際查閱的資料，確保不自行編造。直接引文須核對原文，改寫不得冒充原話；不得虛構出處，或把 repo 未載明的外部原典當作已查證來源。除非使用者另有要求，不必在講稿中加上來源編號、出處註記或額外產生備課來源檔，也不要為了交代來源而反覆插入說明。若 repo 資料不足、來源不明或彼此矛盾，不得自行補寫或當成定論；必要時向使用者簡短說明缺口。
- When creating or editing speech manuscripts, use this opening gratitude wording before introducing the topic: `敬愛的領班、操持講師、以及各位前賢大家晚安!大家好，後學承蒙上天慈悲、祖師弘慈、師恩母德、何老前人開荒台灣的五大犧牲、薛前人的浩然正氣、楊前人的老實修辦、王前人的好好修道、始終如一以及現今李前人慈悲的領導，永續創建一個六合圓滿、紀律化的優質道場，這些都要感恩過去前輩者的犧牲奉獻，建立了道場的基礎跟道場文化，才有今天這麼好的修道環境，後學也感念地區各位點傳師的提拔以及各位前賢們苦口婆心的成全與鼓勵，今天後學才有這個機會可以在這裡學習，後學今天要來學習的題目是OOOO！`
- When creating or editing speech manuscripts, use this closing wording at the end: `後學今日的學習講述到此告一段落，因為後學才疏學淺，還不能夠將道理講述得非常清楚，講述的過程中如有過失懇求上天慈悲，也懇請領班、操持講師以及各位前賢給予後學指正，祝福大家法喜充滿、聖凡如意，謝謝！`
- When moving into each major heading or subheading in a speech manuscript, first add a brief spoken transition that bridges from the preceding topic into the next section, so the new heading is prepared by the prior content rather than appearing abruptly.
- 產生或修改提示稿時，除非使用者明確要求其他格式，須保留原講稿的標題層級與段落順序，將每個非標題自然段落各自濃縮成一句簡潔、完整的提示句，讓講者能據此用自己的話還原該段內容；開場感恩、故事、轉場與結語也須逐段對應，不得合併多段、漏段，或改成按主題重組的提示卡。
- 提示稿每句提示各占一行，行首一律使用 `- ` 作為條列；不用流水號、額外標籤或箭頭串接多個片語，原講稿標題本身的編號則照原樣保留。全文不留任何空白行，包括條列之間及標題前後。
- 提示稿完成後，須核對條列提示句數等於原講稿非標題自然段落數，且逐段順序一致；大題與小題的數量、順序、意義及層級須與對應大綱、講稿一致，並確認每條只有一句話、沒有流水號或額外標籤、全文沒有空白行。
- When creating or editing outlines, follow `docs/模板與提示詞/大綱整理規則.md`.
- For any new speech topic, start by creating or refining the outline first. Confirm the content structure in the outline before expanding it into the speech manuscript.
- Speech manuscripts are expanded versions of their corresponding outlines. Treat each outline/manuscript pair as linked documents: when updating one, immediately review the other and apply any needed synchronized changes.
- For linked outline/manuscript pairs, the speech manuscript must include the outline's major headings and subheadings so the two documents can be read in parallel during preparation and delivery.
- For linked outline/manuscript pairs, keep both the major heading structure and subheading structure synchronized: the number, order, and meaning of corresponding headings should match unless there is an explicit reason to diverge. If the structures intentionally differ, state the reason clearly in the response.
- When revising either an outline or its paired speech manuscript, review whether the corresponding headings in the paired file also need to be updated, and apply synchronized changes immediately when needed.
- After editing any outline or speech manuscript, explicitly inspect the paired file before finishing. Verify the major heading count, subheading count, order, and meaning in both files. Do not assume no change is needed from memory; either synchronize the paired file or state the concrete reason for an intentional difference in the response.
