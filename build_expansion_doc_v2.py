from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def add_title(document: Document, text: str, subtitle: str | None = None) -> None:
	title = document.add_paragraph()
	title_run = title.add_run(text)
	title_run.bold = True
	title_run.font.size = Pt(20)
	title.alignment = WD_ALIGN_PARAGRAPH.CENTER
	if subtitle:
		sub = document.add_paragraph(subtitle)
		sub.alignment = WD_ALIGN_PARAGRAPH.CENTER


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


def add_table(document: Document, headers: list[str], rows: list[list[str]], style: str = 'Light List Accent 1') -> None:
	table = document.add_table(rows=1, cols=len(headers))
	try:
		table.style = style
	except Exception:
		pass
	hdr_cells = table.rows[0].cells
	for idx, h in enumerate(headers):
		hdr_cells[idx].text = h
	for row in rows:
		cells = table.add_row().cells
		for idx, val in enumerate(row):
			cells[idx].text = str(val)


def build_document(path: str) -> None:
	doc = Document()

	# Cover
	add_title(doc, '運務系統功能擴充案（草案）', '依原建置案格式對齊，整理並歸納本次擴充需求')

	# Revision history
	add_heading(doc, '修訂紀錄', 2)
	add_table(doc, ['版本', '日期', '說明', '編寫/覆核'], [
		['v0.1', 'YYYY-MM-DD', '初稿（彙整需求、建置 1.4 時程與成本）', '系統整合/PM'],
	])

	# Executive summary
	add_heading(doc, '1. 專案摘要（Executive Summary）', 1)
	add_bullets(doc, [
		'目標：導入進場證件電子化、進出場掃碼與 OMS 同步、危害告知電子化與簽認追蹤。',
		'效益：縮短人工核對時間、降低誤入風險、提升稽核可追溯性（Excel 匯出、簽認紀錄）。',
		'範疇：車務進場管制、行控擴充、危害告知；介接 OMS，維持既有作業不中斷。',
	])

	# Scope
	add_heading(doc, '2. 專案範疇（Scope）', 1)
	add_table(doc, ['分類', '內容'], [
		['In Scope', '員工/甲乙證改為感應卡與條碼；掃碼帶出工單與人員資訊；OMS 同步進出場；危害告知電子化簽名；簽認紀錄查詢；Excel 匯出。'],
		['Out of Scope', '硬體採購以外之機房改裝；非 OMS 相關他系統功能；第三方授權費用；大型報表平台改造。'],
	])

	# Functional requirements
	add_heading(doc, '3. 需求', 1)
	add_heading(doc, '3.1 功能性需求（Functional Requirements）', 2)
	func_headers = ['模組', '功能', '類型', '說明', '影響/關聯', '驗收重點']
	func_rows = [
		['進場管制', '識別證/甲乙證電子化', '新增', '識別證改為感應卡，旁新增條碼作為身分識別。', 'HR/門禁整合（僅讀取卡號與條碼）', '感應/掃碼能帶出對應人員且通過有效期判斷'],
		['進場管制', '掃工單條碼帶出資訊', '新增', '掃工單條碼，自動跳出工單資訊。', 'OMS 工單資料讀取', '掃碼 1 秒內顯示工單資訊'],
		['進場管制', '進場證件有效性判讀', '新增', '掃描進場證件，自動判讀有效性，通過後帶入進場資訊至 OMS 並同步更新。', 'OMS 寫入進場資訊', '有效/失效規則可設定，寫入 OMS 成功後回饋狀態'],
		['進場管制', '車務進場申請區域調整', '優化', '移除「機廠軌道區」「重要機房區」兩進場區域。', '車務進場申請模組', '既有申請單顯示與選單同步調整且不影響既有資料'],
		['進場管制', '授權碼進場分佈圖（機廠）', '新增', '與行控核發授權碼作業連動，於圖面呈現授權碼、防護措施（SCD/ETS/EBW）、斷電、工程車路徑。', '行控授權碼模組', '圖面定位正確、與授權碼資料一致'],
		['進場管制', '授權碼進場分佈圖（正線）', '新增', '新增 A22、A23 站點至分布圖，呈現授權碼與標示。', '行控授權碼模組', 'A22/A23 顯示正確且可與他站一致操作'],
		['行控', '行控進場申請/審核（機廠/綠線）', '新增', '參照 toc050103/050107 設計；機廠與綠線之區域依需求調整；多日進場未進場日期可修改。', '行控中心', '審核流程可追蹤、修改限制符合規則'],
		['行控', '核發授權碼（機廠/綠線）', '新增', '參照 toc050203 設計，含分佈圖連動。', '行控中心', '授權碼生成、查詢、連動無誤'],
		['行控', '進場率統計', '新增', '新增「機廠」「綠線」統計群組。', '報表', '統計口徑清楚、數值與原始資料一致'],
		['行控', '機廠調車申請/查詢/審核', '新增', '參照 toc040301~040303 設計（待確認細節）。', '車輛調度', '表單欄位與流程符合規範'],
		['行控', '席位擴充與交接表', '新增/優化', '新增 2 個 DC 席位；營運前檢查（occ0201）與席位交接表（occ0202），支援搜尋與共用內容。', '行控作業', '席位可選/查、交接內容共用且權限正確'],
		['行控', '斷復電管理/設定', '優化', 'TSS 參數新增機廠區域供電設備。', '設備參數', '新增參數可維護並套用'],
		['行控', '報修事件與 MMS 介接', '優化', '配合新版 MMS 更新（API 介接待確認）。', 'MMS', '介接測試資料可流通，失敗有重試/告警'],
		['行控', '監控資訊管理版面優化', '優化', '網址加小標題超連結，簡化文字。', 'UI/UX', '版面一致性、操作便捷'],
		['行控', '航班資訊代碼對照表', '新增', '新增代碼對照表供自行維護。', '資料維護', '維護介面與驗證齊備'],
		['行控', '發車提醒文字語音播報', '新增', '新增文字轉語音播報功能。', '通知模組', '語音內容正確、延遲可接受'],
		['行控', '故障代碼管理新增欄位', '優化', '新增「機廠發車與否」欄位。', '代碼維護', '欄位可查可改，影響範圍已測'],
		['行控', '收/發/調車紀錄、調度查詢、列車狀況、駐車位置', '新增', '參照 toc040401/040402/040403/040404/040405/040703 設計（細節待討論）。', '車輛調度', '表單與查詢條件符合設計'],
		['危害告知', '危害告知電子化與簽名', '新增', '掃工單與進場證件後，跳出危害告知條碼；現場掃碼帶出表單並於手機或掃證簽名；後台可查簽認紀錄。', '簽名/稽核', '簽名圖檔與時間人員可追溯、紀錄可查'],
		['共同', '進出場資料匯出 Excel', '新增', '進出場資訊可匯出 Excel。', '報表', '匯出欄位齊全、編碼格式正確'],
	]
	add_table(doc, func_headers, func_rows)

	# Non-functional
	add_heading(doc, '3.2 非功能性需求（Non-Functional Requirements）', 2)
	add_table(doc, ['類別', '要求'], [
		['效能', '掃碼至畫面回應 ≤ 1 秒；授權碼分佈圖載入 ≤ 3 秒（同站點資料）。'],
		['可用性', '服務可用度 99.5%（營運時間）；異常有告警與降級機制。'],
		['資安', '原始碼掃描與弱點掃描通過；存取權限依角色控管；簽名與個資加密儲存。'],
		['稽核', '所有進/離場、簽名、授權碼操作留存稽核紀錄（含人員/時間/IP）。'],
		['相容性', '與現行 OMS 版號相容；不破壞既有作業流程。'],
	])

	# Interfaces
	add_heading(doc, '4. 系統整合與介面', 1)
	add_heading(doc, '4.1 OMS 介面（讀/寫）', 2)
	add_table(doc, ['方向', '資料項目（示意）', '觸發時機'], [
		['讀取', '工單(單號、區域、日期)、人員(工號、姓名、有效期)', '掃工單條碼/掃識別證'],
		['寫入', '進場紀錄(人員、工單、時間、區域、有效性結果)', '進場確認'],
		['寫入', '離場紀錄(人員、工單、時間、結果)', '離場確認'],
	])
	add_heading(doc, '4.2 資料欄位（進出場核心）', 2)
	add_table(doc, ['欄位', '說明'], [
		['PersonId', '人員識別（與 OMS 對應）'],
		['BadgeId / Barcode', '識別證卡號 / 條碼'],
		['WorkOrderId / Barcode', '工單單號 / 條碼'],
		['Area / Station', '進場區域/站點（含機廠/綠線）'],
		['ValidFrom / ValidTo', '證件有效期限'],
		['CheckInAt / CheckOutAt', '進場/離場時間戳'],
		['HazardAck', '危害告知簽認狀態與簽名檔路徑'],
		['AuditTrailId', '稽核紀錄索引'],
	])

	# Workflows
	add_heading(doc, '5. 作業流程（Workflow）', 1)
	add_heading(doc, '5.1 進場', 2)
	add_numbered(doc, [
		'掃工單條碼 → 顯示工單資訊',
		'掃識別證/甲乙證 → 有效性判讀',
		'通過後：寫入 OMS 進場紀錄並回顯成功狀態',
		'（若需）顯示授權碼與分佈圖定位',
	])
	add_heading(doc, '5.2 離場', 2)
	add_numbered(doc, [
		'掃識別證 → 顯示人員當日進場狀態',
		'確認離場 → 寫入 OMS 離場紀錄 → 匯出報表（可選）',
	])
	add_heading(doc, '5.3 現場危害告知', 2)
	add_numbered(doc, [
		'掃工單條碼＋掃進場證件 → 系統跳出危害告知條碼與內容',
		'現場掃條碼帶出表單 → 手機/掃證裝置簽名',
		'簽認完成 → 後台可查簽名與時戳、人員、工單',
	])

	# Acceptance
	add_heading(doc, '6. 測試與驗收', 1)
	add_bullets(doc, [
		'功能驗收：逐項對照 3.1 表，完成操作與資料驗證。',
		'整合驗收：OMS 讀/寫成功率 99.9%，失敗有重試與錯誤記錄。',
		'效能驗收：依 3.2 效能目標抽樣 30 次以上，達標為合格。',
		'資安驗收：原始碼/弱掃零高風險，重大弱點修補完畢。',
		'報表驗收：Excel 匯出欄位與內容與畫面一致。',
	])

	# Deployment
	add_heading(doc, '7. 部署與回復計畫', 1)
	add_numbered(doc, [
		'建立新模組 Feature Flag，先灰度開啟；',
		'資料庫備份與回復腳本完善；',
		'部署後 24 小時監控重點：掃碼延遲、OMS 錯誤率、簽名存取；',
		'回復機制：關閉 Flag、回滾套件、恢復資料。',
	])

	# Risks
	add_heading(doc, '8. 風險與因應', 1)
	add_table(doc, ['風險', '影響', '因應措施'], [
		['OMS 介面不穩', '資料不同步/延遲', '快取與重試、錯誤告警、脫機緩存機制'],
		['掃碼硬體相容性', '現場無法掃描', '選定相容列表、場測與備援流程'],
		['個資/簽名存取', '資安疑慮', '加密保存、最小權限、存取稽核'],
	])

	# Schedule 1.4
	add_heading(doc, '9. 專案時程（對齊 1.4）', 1)
	add_heading(doc, '9.1 預計時程', 2)
	doc.add_paragraph('本專案預計時程為決標日次日起 150/560 日曆天內完成（機關審查期間不計入履約期間）。')
	add_heading(doc, '9.2 各階段時程概述', 2)
	p = doc.add_paragraph()
	r = p.add_run('第一階段：履約期間為決標日次日起 85/320 天（不含審查期間）。')
	r.bold = True
	add_numbered(doc, [
		'需求分析教育訓練、需求訪談及系統分析、系統設計（工作計畫書審查通過次日起 90 日內完成）',
		'軟體開發、系統測試與操作教育訓練、系統測試（第一階段程式設計書審查通過、機關通知起 150 日內完成）',
		'第一階段功能上線導入（系統測試報告審查通過、機關通知起 60 日內完成）',
	])
	add_bullets(doc, [
		'完成第一階段應交付：表 3.6-1 與 3.6-2 項目。',
		'查驗合格後支付契約價金 50%。',
	])
	p = doc.add_paragraph()
	r = p.add_run('第二階段：履約期間為機關通知日次日起 240 天（不含審查期間）。')
	r.bold = True
	add_numbered(doc, [
		'需求訪談及系統分析、系統設計（第一階段查驗通過、機關通知起 60 日內完成）',
		'軟體開發、系統測試與操作教育訓練、系統測試（第二階段程式設計書審查通過、機關通知起 120 日內完成）',
		'第二階段功能上線導入（系統測試報告審查通過、機關通知起 60 日內完成）',
	])
	add_bullets(doc, [
		'完成第二階段應交付：表 3.6-1 與 3.6-2 項目。',
		'竣工驗收合格後支付契約價金 50%。',
		'全案驗收日起提供一年保固維運。',
	])

	# Cost
	add_heading(doc, '10. 成本與工時', 1)
	add_heading(doc, '10.1 報價項目與金額（NTD）', 2)
	add_table(doc, ['工作項目', '金額'], [
		['新增進場管制相關功能', '—'],
		['新增行控相關功能', '—'],
		['資安檢測及修補作業（原始碼掃描、弱點掃描）', '—'],
		['文件製作、交通、公費及管理等', '—'],
	])
	doc.add_paragraph('小計（未稅）：1,498,000')
	doc.add_paragraph('總金額（含稅5%）：1,572,900')
	add_heading(doc, '10.2 工時分析', 2)
	add_table(doc, ['工作項目', '參與人力', '單價', '單位', '數量', '小計'], [
		['專案管理、需求訪談、客戶服務', '專案經理', '—', '式', '1', '448,000'],
		['前端介面設計及調整作業', '前端設計師', '4,500', '人日', '20', '90,000'],
		['功能開發、程式設計', '系統工程師、程式設計師、資安督導', '5,800', '人日', '130', '754,000'],
		['資安檢測與修補、文件製作、交通、公費及管理費', '全員', '—', '式', '1', '206,000'],
	])
	add_bullets(doc, ['薪酬、公費及管理比例依「中華民國資訊軟體協會 資訊服務委外經費估算原則」。'])

	# Change log vs original
	add_heading(doc, '11. 與原建置案差異（Change Log）', 1)
	add_table(doc, ['模組', '項目', '變更類型', '備註'], [
		['進場管制', '識別證電子化＋條碼', '新增', '新增流程與欄位'],
		['進場管制', '工單掃碼自動帶出', '新增', '新增條碼掃描'],
		['進場管制', '進場有效性判讀→OMS 同步', '新增', '新增寫入介面'],
		['進場管制', '進場區域移除（機廠軌道區/重要機房區）', '優化/調整', '移除兩區域選項'],
		['行控', '進場申請/審核（機廠/綠線）', '新增', '與既有流程一致化'],
		['行控', '核發授權碼（機廠/綠線）＋分佈圖', '新增', '呈現授權碼/防護措施/路徑'],
		['行控', '進場率統計（機廠/綠線）', '新增', '新增統計群組'],
		['行控', '席位擴充與交接表', '新增/優化', '新增 2 席位與共用資料庫'],
		['行控', '斷復電管理新增 TSS 參數', '優化', '新增機廠供電設備'],
		['行控', 'MMS 介接', '優化', '待 API 討論'],
		['危害告知', '電子化簽名與追蹤', '新增', '可查簽認紀錄'],
		['共同', 'Excel 匯出', '新增', '進出場資訊匯出'],
	])

	# Appendices
	add_heading(doc, '附錄 A：需求來源整理（摘錄）', 1)
	add_bullets(doc, [
		'進場證件電子化與條碼識別、掃碼帶出工單與人員、OMS 同步。',
		'離場掃識別證更新至 OMS、進出場資訊匯出 Excel。',
		'危害告知電子化：掃碼帶出表單與簽名、簽認紀錄可查。',
		'行控與機廠各模組之功能補強與新表單設計（toc 參考）。',
	])

	doc.save(path)


if __name__ == '__main__':
	build_document('/workspace/expansion_proposal_v2.docx')