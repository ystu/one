# 通用海報 GPT 協作 Prompt

> 使用原則：海報內容每次都會變，所以 prompt 應該先建立「工作流程」和「可替換欄位」，不要綁死在單一活動。每次先貼活動資料，請 GPT 抽出內容，再做無文字底圖；底圖確認後，才進入可編輯文字排版。

## 0. 使用方式

每次做新海報時，先把下面三段貼給 GPT：

1. 「給 GPT 的總指令」
2. 「活動資料」
3. 「你想要的視覺方向」，如果還沒有方向，就請 GPT 先提案

如果活動資料來自 Markdown、Word、會議紀錄或簡章，直接貼原文也可以。請 GPT 先抽資訊，不要一開始就生成圖片。

## 1. 給 GPT 的總指令

```text
請你擔任海報視覺設計協作夥伴，協助我把活動資料整理成可用於 AI 圖像生成與後續排版的海報製作流程。

請遵守以下規則：

1. 先閱讀我提供的活動資料，抽出海報必備資訊。
2. 不要擅自改寫日期、時間、地點、人名、費用、電話、報名方式、QR code 說明或主辦單位。
3. 如果資料不足，請列出缺少項目，並用合理預設繼續。
4. 第一階段只做「無文字背景／底圖」prompt。
5. 底圖 prompt 必須明確要求：no text, no letters, no numbers, no logo, no watermark。
6. 底圖要保留標題區與活動資訊區的留白，方便之後放可編輯文字。
7. 不要把完整中文資訊直接生成在圖片中；繁體中文、日期、人名與電話之後用可編輯文字層處理。
8. 等我確認底圖方向後，再協助規劃文字排版、字級、位置、顏色、遮罩與輸出檢查。

請先輸出：

- 活動資訊摘要
- 海報目標與受眾
- 建議尺寸
- 文字階層
- 2 到 4 個無文字底圖視覺方向
- 每個方向的圖像生成 prompt
- 每個方向的風險或注意事項
```

## 2. 活動資料填空區

把活動資訊貼到這裡，再一起丟給 GPT：

```text
【活動資料】

活動名稱：
主題／副標題：
日期與時間：
報到時間：
地點：
地址：
對象：
費用：
報名方式：
報名截止：
主辦／承辦／協辦：
聯絡人／電話／Email：
QR code 或報名連結文字：
重要注意事項：
其他必放文字：

【活動說明原文】
貼上簡章、Markdown、會議紀錄、課程表或活動說明。
```

## 3. 視覺方向填空區

如果你已經知道方向，貼這段：

```text
【視覺方向】

用途：例如 A4 列印、LINE 分享、Instagram 貼文、限動、簡報首頁
尺寸：例如 A4 直式、1080x1350、1080x1920、1920x1080
調性：例如 溫暖、莊重、青春、現代、喜慶、學術、親切、清爽、高級
主要受眾：
希望出現的元素：
不希望出現的元素：
品牌色／指定色：
必用圖片或 logo：
印刷需求：例如 是否出血、是否需要高解析
```

如果你還沒有方向，貼這段：

```text
我還沒有確定視覺方向。請根據活動資料提出 3 個不同方向，每個方向要包含：

1. 視覺概念
2. 畫面構圖
3. 色彩與光線
4. 適合放文字的位置
5. 無文字底圖 prompt
6. 可能風險
```

## 4. 通用無文字底圖 Prompt 模板

可以請 GPT 依活動替換中括號內容：

```text
[SIZE AND ORIENTATION] poster background, no text, no letters, no numbers, no logo, no watermark.

Create a [VISUAL STYLE] background for [EVENT TYPE] about [CORE THEME].
The mood should feel [MOOD KEYWORDS].
The audience is [AUDIENCE], so the image should feel [AUDIENCE-APPROPRIATE TONE].

Main visual: [MAIN SUBJECT OR SCENE].
Secondary details: [SUPPORTING ELEMENTS].
Environment: [PLACE / SEASON / CULTURAL CONTEXT].
Lighting: [LIGHTING DIRECTION].
Color palette: [COLOR DIRECTION].

Composition requirements:
- Leave clean negative space at [TITLE AREA] for the main title.
- Leave a calm readable area at [DETAIL AREA] for date, venue, CTA, and organizer.
- Avoid clutter behind future text areas.
- Keep important subjects away from page edges and safe margins.

Quality direction:
polished poster background, high detail, balanced composition, refined lighting, suitable for editable typography overlay, print-friendly, [ASPECT RATIO].

Avoid:
text, letters, numbers, readable signs, logos, watermarks, distorted hands, awkward faces, extra limbs, messy crowding, unrelated symbols, overly commercial stock-photo feeling.
```

## 5. GPT 修 Prompt 回合用語

當 GPT 的底圖 prompt 太散、太像廣告圖、留白不夠，或容易生成錯字時，貼這段：

```text
請保留目前方向，但把 prompt 修成更適合正式海報底圖：

- 不要任何文字、字母、數字、招牌、logo、浮水印
- 上方或主標題位置要保留乾淨留白
- 下方或資訊區要保留可讀性好的空間
- 主體不要壓到安全邊界
- 風格要符合活動受眾與活動性質
- 不要太商業廣告感，不要廉價素材感
- 請讓畫面更有層次、光線更精緻、構圖更穩定
- 請輸出一版可以直接丟給圖像生成工具的英文 prompt
```

如果生成圖已經有明顯問題，貼這段：

```text
這版有以下問題：

- 問題一：
- 問題二：
- 問題三：

請不要改變我喜歡的部分：

- 保留一：
- 保留二：

請根據問題重寫 prompt，目標是生成下一版無文字底圖。不要加入任何活動文字。
```

## 6. 底圖確認後的排版協作 Prompt

底圖方向確認後，再貼這段。這時不要重新生成底圖。

```text
我已確認海報底圖。請不要重新生成或改動底圖構圖、光線、人物、場景。現在只協助規劃可編輯文字排版。

請依照我提供的活動資料，使用繁體中文設計文字層。所有日期、時間、地點、人名、費用、電話、Email、報名方式、主辦單位必須完全照原文，不得自行改寫。

請輸出：

1. 文字資訊分層：主標題、副標題、關鍵資訊、補充資訊、CTA、主辦單位
2. 每個文字區塊的位置
3. 建議字級比例，而不是只給固定字級
4. 字體風格建議
5. 顏色與對比建議
6. 是否需要半透明遮罩、漸層遮罩、線條、色塊或資訊框
7. 手機觀看時最先讀到的三個資訊
8. 列印時的安全邊界與可讀性檢查

請另外列出「不能改錯的原始資料」，方便我最後核對。
```

## 7. 文字層內容模板

排版前可先要求 GPT 整理成這種結構：

```text
【主標題】

【副標題／主題句】

【關鍵資訊】
- 日期：
- 時間：
- 地點：
- 地址：
- 對象：
- 費用：
- 報名截止：

【CTA】
- 報名方式：
- QR code 說明：
- 聯絡方式：

【主辦與備註】
- 主辦：
- 承辦：
- 協辦：
- 注意事項：
```

## 8. 最終檢查 Prompt

```text
請幫我檢查這張活動海報是否可以交付。請用活動原始資料逐項核對，不要憑印象。

請特別檢查：

- 活動名稱是否完整
- 主題與副標題是否正確
- 日期、時間、報到時間是否正確
- 地點、地址、電話是否正確
- 費用、報名截止、報名方式是否正確
- 主辦、承辦、協辦或聯絡資訊是否正確
- 是否有 AI 生成錯字、亂碼、假文字、假 logo、假 QR code
- 手機螢幕上是否能先讀到活動名稱、日期、地點
- 重要資訊是否都在安全邊界內
- 文字對比是否足夠
- 背景是否干擾文字可讀性

最後請輸出：

1. 必修正問題
2. 可微調建議
3. 可以交付前仍需人工確認的資訊
```
