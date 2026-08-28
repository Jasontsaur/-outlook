# outlook-mcp

本機 MCP Server，讓 Claude Code / Claude Desktop / 本機 Cowork 透過 Outlook 2016
的 COM Automation 讀取、搜尋、分類建議，並在你確認後移動郵件。

## 提供的工具

**唯讀（不會修改任何東西）：**

- `outlook_list_folders`：列出所有帳號/PST 下的所有資料夾與郵件數量
- `outlook_list_recent_emails`：列出某資料夾最新 N 封郵件的中繼資料（主旨/寄件人/時間，不含內文）
- `outlook_search`：依日期範圍/主旨關鍵字/寄件人關鍵字搜尋，回傳 metadata（不含內文）
- `outlook_get_email`：取得單一郵件的完整內文與附件檔名
- `outlook_classify_recent`：搜尋 + 依 `rules.yaml` 分類建議，回傳各分類數量與逐封建議去向，**不會移動任何郵件**

**寫入（有安全防護）：**

- `outlook_create_folder`：建立資料夾（可重複呼叫，已存在就跳過）
- `outlook_preview_move`：**唯讀**，顯示如果真的搬移「會發生什麼事」，同時會寫一份 CSV 稽核紀錄到 `staging/`
- `outlook_move_emails`：真正搬移郵件，**沒有 `confirm=True` 一律拒絕執行**

**安全流程（務必遵守）：**

```
outlook_search / outlook_classify_recent（唯讀，先看整體狀況）
        ↓
outlook_preview_move（唯讀，列出這批信實際會怎麼被搬）
        ↓
使用者親自確認要不要搬
        ↓
outlook_move_emails(confirm=True)（才會真的動手）
```

不要為了「省一次確認」而直接呼叫 `outlook_move_emails`；`outlook_preview_move` 的存在就是為了讓你在真正搬信之前，先看到會發生什麼事。

## 前置需求

- **Windows** 11（或 10），且 **Outlook 2016 Classic** 已安裝並至少開啟登入過一次
- Python 3.10+（COM 呼叫必須跑在能存取這個 Windows 使用者 session 的 Python 上，
  不能在遠端伺服器或容器裡執行）
- 執行這支程式時，建議 Outlook 本身也保持開啟中

## 安裝 / 更新

```powershell
cd C:\ClaudeTools\outlook-repo\outlook-mcp
git pull origin claude/outlook-email-organization-qjyop2
.venv\Scripts\activate
pip install -r requirements.txt
```

## Step 1：先跑最基本的連線測試（不透過 MCP）

```powershell
python test_connection.py
```

預期會印出你的資料夾清單，以及 Inbox 最新 10 封信的主旨/寄件人/時間。
**這一步一定要先成功，再進到下一步接 MCP** — 如果這裡就失敗，接 MCP 也不會成功。

### 常見失敗排解

| 錯誤現象 | 可能原因 / 解法 |
| --- | --- |
| `Could not connect to Outlook via COM` | Outlook 沒開過、不是同一個 Windows 使用者 session、或 Outlook 正在跳出對話框（例如憑證提示）擋住 |
| `pywintypes.com_error` 授權/安全性警告 | 部分企業環境的 Outlook 安全性設定會擋第三方程式存取，需請 IT 調整「程式化用戶端存取」設定 |
| 找不到資料夾 (`Folder not found`) | Inbox/Sent Items 等預設資料夾已透過 `GetDefaultFolder` 處理語系問題；自訂資料夾（如 `Projects/XX工程`）要用 Outlook 裡實際看到的名稱 |

## Step 2：接上 MCP

**Claude Code：**

```powershell
claude mcp add outlook -- C:\ClaudeTools\outlook-repo\outlook-mcp\.venv\Scripts\python.exe C:\ClaudeTools\outlook-repo\outlook-mcp\server.py
```

**Claude Desktop App：** 編輯 `%APPDATA%\Claude\claude_desktop_config.json`：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\ClaudeTools\\outlook-repo\\outlook-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\ClaudeTools\\outlook-repo\\outlook-mcp\\server.py"]
    }
  }
}
```

存檔後完整重啟該用戶端。

## 分類規則

`rules.yaml` 是 Claude 判斷「這封信該分去哪」的依據，`OUTLOOK_EMAIL_RULES.md`
是同一份規則的人類可讀版本。目前已啟用的規則：

- 寄件者網域含 `ruentex` → `Internal`

其餘（工程專案/客戶/供應商/財務/通知）先以 `<TBD>` 佔位、`enabled: false`
停用，等你補上實際的專案代號、客戶/供應商名稱或網域，把 `enabled` 改成
`true` 即可生效。

## 建議的使用流程

```
1. 跟 Claude 說：「用 outlook_classify_recent 掃過去一年的 Inbox，
   不要搬動任何東西，先給我看分類統計」
        ↓
2. 檢查 counts_by_destination 跟低信心/UNMATCHED 的候選信件，
   覺得規則不準就調整 rules.yaml，重跑一次
        ↓
3. 規則穩定後，針對某個分類跟 Claude 說：
   「這批 Projects/XX工程 的信，先幫我 preview_move」
        ↓
4. 看過 outlook_preview_move 的結果、確認沒問題後才說：
   「確認，幫我搬」— Claude 才會呼叫 outlook_move_emails(confirm=True)
```

## 目錄結構

```
outlook-mcp/
├── .venv/                    # 虛擬環境（不進版控）
├── server.py                 # MCP server 進入點，定義所有 tools
├── outlook.py                # Outlook COM 存取邏輯（讀取 + 受保護的寫入）
├── classify.py                # 規則比對邏輯
├── rules.yaml                 # 機器可讀的分類規則
├── OUTLOOK_EMAIL_RULES.md      # 人類可讀的分類規則說明
├── test_connection.py          # 不經 MCP 的獨立連線測試
├── staging/                    # outlook_preview_move 產生的 CSV 稽核紀錄（不進版控）
├── requirements.txt
└── README.md
```
