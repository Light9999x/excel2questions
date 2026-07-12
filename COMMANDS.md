# 網站手動開關指令

請在此資料夾執行：

```powershell
cd C:\Users\963er\MyStuff\ProjectsWithAI\environmental-knowledge-website
```

## 開啟網站

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\start-site.ps1
```

開啟後用瀏覽器進入：

```text
http://127.0.0.1:5500/index.html
```

## 關閉網站

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\stop-site.ps1
```

## 查看目前狀態

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\status-site.ps1
```

## 備註

- 這個網站是本機網站，只會綁定在 `127.0.0.1:5500`。
- 開啟網站時會在背景啟動 Python 靜態伺服器。
- 關閉網站時會優先使用 `.site-server.pid`，如果檔案不存在，也會嘗試關閉佔用 `5500` 的 Python 網站程序。
