# outlook-mcp (V0.1 - read-only)

本機 MCP Server，讓 Claude Code / Claude Desktop / 本機 Cowork 透過 Outlook 2016
的 COM Automation 讀取郵件資料。**這個版本完全唯讀**，只提供：

- `outlook_list_folders`：列出所有帳號/PST 下的所有資料夾與郵件數量
- `outlook_list_recent_emails`：列出某資料夾最新 N 封郵件的中繼資料（主旨/寄件人/時間，不含內文）

沒有任何搜尋、分類、移動、刪除功能 — 這些會在確認 V0.1 可正常運作後，於 V0.2～V0.4 陸續加入。

## 前置需求

- **Windows** 11（或 10），且 **Outlook 2016 Classic** 已安裝並至少開啟登入過一次
- Python 3.10+（COM 呼叫必須跑在能存取這個 Windows 使用者 session 的 Python 上，
  不能在遠端伺服器或容器裡執行）
- 執行這支程式時，建議 Outlook 本身也保持開啟中

## 安裝

```powershell
cd C:\ClaudeTools\outlook-mcp
python -m venv .venv
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
| 找不到資料夾 (`Folder not found`) | 資料夾路徑要用實際名稱組成，例如多帳號時可能是 `你的Email/Inbox` 而非單純 `Inbox` |

## Step 2：接上 MCP，讓 Claude Code 呼叫

在 Claude Code 的 MCP 設定（例如專案的 `.mcp.json`，或全域 `claude mcp add`）加入：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\ClaudeTools\\outlook-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\ClaudeTools\\outlook-mcp\\server.py"]
    }
  }
}
```

或用 CLI：

```powershell
claude mcp add outlook -- C:\ClaudeTools\outlook-mcp\.venv\Scripts\python.exe C:\ClaudeTools\outlook-mcp\server.py
```

重新啟動 Claude Code 後，跟它說：

```
用 outlook_list_folders 列出我的 Outlook 資料夾，
再用 outlook_list_recent_emails 看 Inbox 最新 10 封信。
```

確認能正確回傳資料，V0.1 就算完成。

## 目錄結構

```
outlook-mcp/
├── .venv/                 # 虛擬環境（不進版控）
├── server.py              # MCP server 進入點，定義 tools
├── outlook.py             # Outlook COM 存取邏輯（純讀取）
├── test_connection.py     # 不經 MCP 的獨立連線測試
├── requirements.txt
└── README.md
```

## 下一步（V0.2 起，尚未實作）

- `outlook_search`：依日期/寄件人/主旨/內文關鍵字搜尋，先回傳 metadata 而非完整內文
- `outlook_get_email`：只在確定要看某封信時才取得完整內文與附件資訊
- 之後才會加入 `outlook_preview_move` / `outlook_move_email` 等有寫入行為的工具，
  且一定會先有 Preview + 使用者確認機制，才會真正執行移動。
