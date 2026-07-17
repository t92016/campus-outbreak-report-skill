# 校園集體症狀通報系統 — 建置技能（Skill）

這是一套從**新北市忠義國小校園出現集體腸胃不適症狀事件通報系統**（2026-07 開發完成並正式
上線：https://t92016.github.io/ ）萃取出來的標準化建置流程，讓任何電腦上的任何 AI 助理，
只要讀完本 repo 的文件，就能協助使用者重建一套相同架構的通報系統。

## 給 AI 助理

請先讀 [`SKILL.md`](./SKILL.md)，裡面定義了完整的觸發流程與詢問順序。

## 給人看的快速導覽

- [`SKILL.md`](./SKILL.md)：AI 助理的行為準則（觸發詞、詢問順序、核心原則）
- [`docs/建置手冊.md`](./docs/建置手冊.md)：完整部署步驟（表單→試算表→GAS Web App→GitHub Pages）
- [`docs/踩坑記錄.md`](./docs/踩坑記錄.md)：11 個實際踩過的坑與解法
- [`docs/欄位設計指南.md`](./docs/欄位設計指南.md)：為什麼要用「動態依標題文字尋找欄位」
- [`reference/Code.gs`](./reference/Code.gs)：可直接執行的參考實作（Apps Script 完整程式碼）
- [`reference/github-pages-index.html`](./reference/github-pages-index.html)：GitHub Pages 首頁外殼參考實作

## 技術架構總覽

```
Google 表單（導師通報）
    ↓
Google 試算表（彙整總表 = 正本資料）
    ↓ Apps Script 自動處理
├─ 各年級/角色分流工作表
└─ 去識別化總覽
    ↓
GAS Web App（doGet 依 ?page= 路由：首頁 / 各角色通報頁 / 管理者後台）
    ↓
GitHub Pages（純外殼，iframe 嵌入 GAS Web App）
```

## 使用方式

對你的 AI 助理說：**「建立校園集體症狀通報系統」**，並讓它讀取這個 repo
（`git clone` 或直接讓它用網路工具讀取 raw 內容），它就會依照 `SKILL.md` 定義的流程，
詢問你表單規格、學校識別資訊、導覽列規劃、密碼設定方式，然後協助你一步步建置完成。

## 授權與使用限制

本 repo 內容（包含參考程式碼）以「盡量重用、避免重造輪子」為目的公開分享，
使用時請自行確保符合貴校的個資保護政策與相關法規。
