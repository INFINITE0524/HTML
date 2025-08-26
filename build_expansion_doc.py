from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def add_title(document: Document, text: str) -> None:
	title = document.add_paragraph()
	title_run = title.add_run(text)
	title_run.bold = True
	title_run.font.size = Pt(20)
	title.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_heading(document: Document, text: str, level: int = 1) -> None:
	document.add_heading(text, level=level)


def add_bullets(document: Document, items: list[str]) -> None:
	for item in items:
		p = document.add_paragraph(style='List Bullet')
		p.add_run(item)


def add_numbered(document: Document, items: list[str]) -> None:
	for item in items:
		p = document.add_paragraph(style='List Number')
		p.add_run(item)


def add_table(document: Document, headers: list[str], rows: list[list[str]]) -> None:
	table = document.add_table(rows=1, cols=len(headers))
	table.style = 'Light List Accent 1'
	hdr_cells = table.rows[0].cells
	for idx, h in enumerate(headers):
		hdr_cells[idx].text = h
	for row in rows:
		cells = table.add_row().cells
		for idx, val in enumerate(row):
			cells[idx].text = str(val)


def build_document(path: str) -> None:
	doc = Document()

	# Title
	add_title(doc, '運務系統功能擴充案（草案）')
	p = doc.add_paragraph('（依原建置案格式對齊，整併本次擴充需求）')
	p.alignment = WD_ALIGN_PARAGRAPH.CENTER

	# 1. 專案背景與目的
	add_heading(doc, '1. 專案背景與目的', 1)
	background = [
		'進場證件電子化：公司員工識別證、甲乙證改為感應卡，識別證旁新增條碼（bar code）作為身分識別。',
		'與 OMS 同步：進出場資訊經掃碼確認後，於 OMS 即時帶入與更新（進場資訊同步更新）。',
		'資料查詢與報表：進出場資訊可匯出 Excel，利於稽核與追蹤。',
	]
	add_bullets(doc, background)

	# 2. 需求概述與責任分工
	add_heading(doc, '2. 需求概述與責任分工', 1)
	add_heading(doc, '2.1 進場資訊確認（硬體/人力資源處）', 2)
	add_bullets(doc, [
		'掃工單條碼（bar code）自動帶出工單資訊。',
		'可掃進場證件並自動判讀有效性，確認後帶入進場資訊至 OMS，並同步更新進場資訊。',
	])
	add_heading(doc, '2.2 離場資訊確認（軟體/數發中心）', 2)
	add_bullets(doc, [
		'掃識別證即時判斷人員是否離場，確認後更新至 OMS。',
		'進出場資訊可匯出為 Excel。',
	])
	add_heading(doc, '2.3 現場危害告知電子化（軟體/數發中心）', 2)
	add_bullets(doc, [
		'掃工單條碼與進場證件後，系統自動跳出對應「危害告知」條碼與內容。',
		'現場人員掃描後帶出表單，於手機或掃證裝置上完成簽名。',
		'後台可查詢人員簽認紀錄以便追蹤。',
	])

	# 3. 報價與工時分析
	add_heading(doc, '3. 報價與工時分析', 1)
	add_heading(doc, '3.1 報價項目與金額（NTD）', 2)
	quote_rows = [
		['新增進場管制相關功能', '—'],
		['新增行控相關功能', '—'],
		['資安檢測及修補作業（原始碼掃描、弱點掃描）', '—'],
		['文件製作、交通、公費及管理等', '—'],
	]
	add_table(doc, ['工作項目', '金額'], quote_rows)
	p = doc.add_paragraph('小計（未稅）：1,498,000')
	p = doc.add_paragraph('總金額（含稅5%）：1,572,900')

	add_heading(doc, '3.2 按工時分析', 2)
	labor_headers = ['工作項目', '參與人力', '單價', '單位', '數量', '小計']
	labor_rows = [
		['專案管理、需求訪談、客戶服務', '專案經理', '—', '式', '1', '448,000'],
		['前端介面設計及調整作業', '前端設計師', '4,500', '人日', '20', '90,000'],
		['功能開發、程式設計', '系統工程師、程式設計師、資安督導', '5,800', '人日', '130', '754,000'],
		['資安檢測及修補作業、文件製作、交通、公費及管理費', '全員', '—', '式', '1', '206,000'],
	]
	add_table(doc, labor_headers, labor_rows)
	doc.add_paragraph('註：薪酬、公費及管理比例依「中華民國資訊軟體協會 資訊服務委外經費估算原則」。')

	# 4. 專案時程（依 1.4 格式）
	add_heading(doc, '4. 專案時程（對齊原建置案 1.4 範本）', 1)
	add_heading(doc, '4.1 預計時程', 2)
	doc.add_paragraph('本專案預計時程為決標日次日起 150/560 日曆天內完成（機關審查期間不計入履約期間）。')

	add_heading(doc, '4.2 各階段時程概述', 2)
	p = doc.add_paragraph()
	r = p.add_run('第一階段：履約期間為決標日次日起 85/320 天（不含審查期間）。')
	r.bold = True
	add_numbered(doc, [
		'需求分析教育訓練、需求訪談及系統分析、系統設計（上述項目須於工作計畫書審查通過次日起 90 日曆天內完成）。',
		'軟體開發、系統測試與操作教育訓練、系統測試（上述項目須於第一階段程式設計書審查通過、機關通知日次日起 150 日曆天內完成）。',
		'第一階段功能上線導入（須於第一階段系統測試報告審查通過、機關通知日次日起 60 日曆天內完成）。',
	])
	add_bullets(doc, [
		'須完成第一階段應交付項目，詳表 3.6-1 及表 3.6-2。',
		'自工作項目(系統測試)啟始日起至本階段查驗合格止，廠商應適當安排專案組織成員協助系統功能測試與上線導入。',
		'須於工作項目(上線導入)到期日前提報第一階段查驗，查驗合格後付款契約價金 50%。',
	])
	p = doc.add_paragraph()
	r = p.add_run('第二階段：履約期間為機關通知日次日起 240 天（不含審查期間）。')
	r.bold = True
	add_numbered(doc, [
		'需求訪談及系統分析、系統設計（須於第一階段查驗通過、機關通知日次日起 60 日曆天內完成）。',
		'軟體開發、系統測試與操作教育訓練、系統測試（須於第二階段程式設計書審查通過、機關通知日次日起 120 日曆天內完成）。',
		'第二階段功能上線導入（須於第二階段系統測試報告審查通過、機關通知日次日起 60 日曆天內完成）。',
	])
	add_bullets(doc, [
		'須完成第二階段應交付項目，詳表 3.6-1 及表 3.6-2。',
		'自工作項目(系統測試)啟始日起至驗收合格止，廠商適當安排專案組織成員協助系統功能測試與上線導入。',
		'須於工作項目(上線導入)到期日前提報竣工，驗收合格後付款契約價金 50%。',
		'全案驗收合格通過日起，提供一年本專案系統保固維運。',
	])

	add_heading(doc, '4.3 各階段完成功能項目一覽表（表 1.4-1）', 2)
	feature_headers = ['階段', '模組', '作業']
	feature_rows = [
		['第一階段', '共同性', '基本資料管理'],
		['第一階段', '共同性', '維護管理'],
		['第一階段', '車務管理', '車務日誌管理'],
		['第一階段', '車務管理', '線上勤務管理'],
		['第一階段', '車務管理', '進場管制作業'],
		['第一階段', '車務管理', '司機員勤務裝備管理'],
		['第一階段', '行控管理', '例行作業'],
		['第一階段', '行控管理', '事件管理'],
		['第一階段', '行控管理', '日誌及報表管理'],
		['第二階段', '車務管理', '列車管制作業'],
		['第二階段', '車務管理', '行政管理'],
		['第二階段', '車務管理', '事故資料管理'],
		['第二階段', '行控管理', '輔助資訊管理'],
		['第二階段', '行控管理', '行政管理'],
		['第二階段', '行控管理', '簡訊管理'],
	]
	add_table(doc, feature_headers, feature_rows)

	# 5. 附件：功能需求說明（整理自 PDF 文字）
	add_heading(doc, '附錄 A：功能需求說明（整理）', 1)
	add_heading(doc, 'A. 與進場管制相關（節錄與歸納）', 2)
	ingress_items = [
		'功能表分類與名稱優化：配合組織異動與使用習慣調整左側頁籤分類/中文名稱（含權限設定）。',
		'車務進場申請作業：進場區域調整（移除「機廠軌道區」「重要機房區」）。',
		'行控進場申請作業（機廠）：參照 toc050103 設計；進場區域含「機廠軌道區」「機廠重要機房區」。',
		'行控進場申請審核作業優化：多日進場工單，未進場之日期可修改。',
		'行控進場申請審核作業（機廠）：參照 toc050107 設計。',
		'行控核發授權碼作業（機廠）：參照 toc050203 設計。',
		'授權碼進場分佈圖（青埔/蘆竹機廠）：與「行控核發授權碼作業（機廠）」連動，呈現授權碼、防護措施（SCD、ETS、EBW）、斷電資料、工程車路徑等。',
		'授權碼進場分佈圖（正線軌道）：新增 A22 老街溪、A23 中壢車站於分布圖；呈現授權碼、防護措施、斷電資料、工程車路徑位置。',
		'進場類型申請時限設定：行控中心新增「機廠」「綠線」設定群組；開放申請日期區間可指定（如春節前兩週）。',
		'行控進場申請作業（綠線）：參照 toc050103 設計，原則不區分正線與機廠（區域類別待討論）。',
		'行控進場申請審核作業（綠線）：參照 toc050107 設計。',
		'行控核發授權碼作業（綠線）：參照 toc050203 設計。',
		'授權碼進場分佈圖（綠線）：參照 toc050206 設計；圖面分為「綠線」「北機廠」。',
		'進場率統計作業：行控中心新增「機廠」「綠線」統計群組。',
		'機廠調車申請/查詢/審核作業：分別參照 toc040301/040302/040303 設計（待需求討論）。',
	]
	add_numbered(doc, ingress_items)

	add_heading(doc, 'B. 與行控相關（節錄與歸納）', 2)
	control_items = [
		'新增 2 個 DC 席位（名稱待定），並新增該席位之營運前檢查（occ0201）與席位交接檢查表（occ0202），可搜尋。',
		'功能分類整理優化：首頁與左側頁籤分類整理（含權限設定）。',
		'斷復電管理／設定：TSS 參數新增機廠區域供電設備。',
		'交接事項：改為資料庫共用交接內容並重新規劃表格；新增 2 個 DC 席位。',
		'報修事件：配合新版維修管理系統（MMS）更新（是否提供 API 介接待討論）。',
		'監控資訊管理：版面優化，網址加入小標題超連結，簡化版面文字。',
		'航班資訊：新增「代碼對照表」，使用者可維護航空公司與目的地代碼中文名稱。',
		'發車提醒：增加文字語音播報功能。',
		'故障代碼管理：新增「機廠發車與否」欄位。',
		'收車/發車/調車紀錄表：分別參考 toc040401、toc040402、toc040403 設計（含 toc040701 代碼參數設定，待討論）。',
		'車輛調度查詢作業：參考 toc040404 設計（待討論）。',
		'列車狀況及紀錄：參考 toc040405 設計（待討論）。',
		'車組駐車位置調整：參考 toc040703 設計（待討論）。',
	]
	add_numbered(doc, control_items)

	# 6. 備註與付款條件
	add_heading(doc, '6. 備註與付款條件', 1)
	add_numbered(doc, [
		'專案工作內容詳附件「功能需求說明」，未盡事宜由雙方另行商議。',
		'參與本專案人員：專案經理、前端設計師、系統工程師、程式設計師、資安督導（共計 5 位）。',
		'付款條件：依契約條款分階段交付驗收，合格後支付該期約定款項。',
		'報價不含系統開發/建置所需硬體環境（設備）及作業系統、資料庫等第三方軟體授權費用。',
	])

	doc.save(path)


if __name__ == '__main__':
	build_document('/workspace/expansion_proposal.docx')