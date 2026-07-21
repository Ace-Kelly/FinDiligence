# FinDiligence · 金融尽调公开信息自动核查工具

批量做「网络核查」时，把**逐站搜公司、截图、命名、归档、贴进 Word** 这套重复劳动交给脚本，
你只负责在浏览器里搜索和过验证码。

一次核查动辄几十家主体 × 几十个网站，几百上千张截图。手工做，光是给图片改名和往 Word 里
贴就要几个小时，还容易贴错位置。这个工具把除「看页面」以外的环节全接管了。

```
打开一个网站 → 依次把所有公司搜一遍并截图 → 换下一个网站
        ↑ 你只做这一步          ↑ 脚本负责：复制搜索词、截图、命名、建目录、生成 Markdown 和 Word
```

**适用场景**：债券发行人存续期月度核查、IPO/并购尽调的公开信息筛查、授信客户背景核查、
反洗钱负面信息筛查（adverse media screening）—— 凡是要「逐个主体、逐个网站、截图留痕」的活。

---

## 目录

- [为什么不做浏览器自动化](#为什么不做浏览器自动化)
- [截图为什么是全屏](#截图为什么是全屏而不是网页截图)
- [环境要求](#环境要求)
- [快速开始](#快速开始)
- [配置一：换成你自己的公司清单](#配置一换成你自己的公司清单)
- [配置二：增删核查网站](#配置二增删核查网站)
- [配置三：自定义文件夹结构](#配置三自定义文件夹结构)
- [命令参考](#命令参考)
- [常见问题](#常见问题)

---

## 为什么不做浏览器自动化

这个项目最初的版本用 Playwright 驱动浏览器自动填搜索框、自动点查询，**后来整个删掉了**。

原因很实际：政府、司法、征信类网站的风控普遍会识别自动化浏览器特征。用 Playwright/Selenium
跑，轻则频繁弹验证码、搜索无结果，重则登录态失效、账号被限制。核查工作要的是**结果可信、
能交付**，不是跑得快 —— 图没截成、账号被封，比手工慢那点更亏。

所以现在的设计是：

- **驱动你自己日常用的浏览器**（Edge / Chrome），保留你已有的登录态和 Cookie
- 脚本**不注入任何自动化特征**，不控制页面，只做三件事：把搜索词复制到剪贴板、在你按回车后
  截图、按规则命名归档
- 验证码、登录、翻页、点进详情页 —— **全部由你本人在浏览器里完成**

本工具不绕过验证码、不做反爬对抗、不批量抓取数据，只是把你本来就要做的截图工作自动归档。

## 截图为什么是「全屏」而不是网页截图

核查底稿通常要求截图能看到 **Windows 任务栏右下角的系统时间**，用来证明核查时点。浏览器内置
的网页截图拍不到任务栏，所以这里用的是系统级全屏截图。

代价是截图瞬间浏览器必须在前台。脚本会在截图前自动最小化终端窗口，免得终端被拍进图里。

### 多显示器要注意

如果你外接了显示器，**不要**直接用 `ImageGrab.grab(all_screens=True)` 这类整屏抓取 —— 它拍的是
整个虚拟桌面，两块屏加中间空隙全在一张图里，有效内容可能只占三分之一，插进 Word 缩到页宽后
根本看不清。

`capture_screen.py` 的做法是：拿到浏览器窗口句柄 → 判断它在哪块显示器 → 只截那一块。
调用方（`run.ps1`）会把浏览器句柄传进来，不靠猜。

另外，**高 DPI 屏（缩放 150%/200%）截出来的文字比普通 1080p 屏清晰得多**，因为页面本身就是按
2 倍像素渲染的。有条件的话把浏览器放在高分屏上截。

## 环境要求

| 依赖 | 用途 | 必需 |
|---|---|:---:|
| Windows 10/11 | 全屏截图和窗口控制用了 Win32 API | 是 |
| PowerShell 7+ | 主脚本（Windows PowerShell 5.1 也能跑，但中文显示更稳的是 7） | 是 |
| Python 3 + Pillow | 截图（`pip install pillow`） | 是 |
| Edge 或 Chrome | 你平时用的那个，保留登录态 | 是 |
| [Pandoc](https://pandoc.org/installing.html) | 把 Markdown 导成 Word | 否 |

不装 Pandoc 也能正常截图和生成 Markdown，只是不会自动导出 `.docx`。

首次运行如果提示「禁止运行脚本」，执行一次：

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

## 快速开始

```powershell
git clone <你的仓库地址>
cd FinDiligence
pip install pillow

# 先拿默认示例试跑（4 个站 × 3 家虚构公司）
.\run.ps1
```

跑起来之后每张图都会停下来等你：

| 按键 | 作用 |
|---|---|
| `回车` | 截图并进入下一张 |
| `r` | 重截当前这张 |
| `o` | 重新打开本站网址 |
| `s` | 跳过这一张 |
| `n` | 跳过本站剩余公司，换下个网站 |
| `q` | 退出（已截的图都保留） |

**随时可以 Ctrl+C 退出**，下次再跑会自动跳过已经存在的截图，从断点继续。

---

# 配置一：换成你自己的公司清单

公司清单支持两种格式，按你的核查粒度选。

## 格式 A：纯文本（`.txt`）—— 每站每家一张图

最简单，一行一个公司名：

```text
# 我的清单.txt
# 以 # 开头的行会被忽略，可以拿来分组备注

# —— 发行人 ——
某某集团有限公司
某某控股股份有限公司

# —— 担保人 ——
某某融资担保有限公司
```

跑起来：

```powershell
.\run.ps1 -Companies 我的清单.txt
```

每个网站会给每家公司截 **1 张图**，文件名就是公司名。适合：授信客户背景核查、
反洗钱筛查这类「一家公司一条记录」的场景。

## 格式 B：结构化（`.json`）—— 一家公司多个核查对象

债券尽调这类场景基本都要用它：同样是裁判文书网，一家发行人要把**自己、几家重要子公司、
十几位董监高**挨个搜一遍，每人一张图；而百度可能只需要搜发行人本身。

```jsonc
{
  "示例集团": {                          // ← 内部标识，随便起，不出现在文件名里
    "name": "示例集团有限公司",           // ← 公司显示名，也是文件夹名
    "group": "发行人",                    // ← 可选，会多建一层分组目录
    "sites": {
      "03": {                             // ← 对应 sites.json 里的网站 id
        "folderName": "03-裁判文书",       // ← 这个站在这家公司下的文件夹名
        "paragraphs": [                   // ← 这个站要截的每一张图
          {
            "text": "发行人：示例集团有限公司",           // 写进 Markdown/Word 的标题行
            "imgName": "发行人：示例集团有限公司.jpg",     // 截图文件名
            "searchName": "示例集团有限公司"              // 复制到剪贴板的搜索词
          },
          {
            "text": "董事长：张三",
            "imgName": "董事长：张三.jpg",
            "searchName": "张三"
          }
        ]
      },
      "04": {
        "folderName": "04-百度",
        "paragraphs": [
          { "text": "发行人：示例集团有限公司",
            "imgName": "发行人：示例集团有限公司.jpg",
            "searchName": "示例集团有限公司" }
        ]
      }
    }
  }
}
```

三个字段的分工是这个格式的精髓：

| 字段 | 去哪了 | 举例 |
|---|---|---|
| `text` | Word 里图片上方的标题行 | `发行人：示例集团有限公司` |
| `imgName` | 磁盘上的文件名 | `发行人：示例集团有限公司.jpg` |
| `searchName` | 自动复制到剪贴板，你直接 Ctrl+V | `示例集团有限公司` |

**为什么要分开**：Word 底稿里习惯写「发行人：某某集团」，但搜索框里只能搜「某某集团」。
分开写，标题好看、搜索也准。

完整可运行的例子见 [`config/companies.example.json`](config/companies.example.json)。

## 偷懒技巧：让 `text` 和 `searchName` 自动对齐

如果你的 `searchName` 只是 `text` 去掉前缀，可以不写 `searchName`，改在网站配置里加
`stripPrefixes`，脚本复制时会自动剥掉前缀和 `1、` 这类编号：

```jsonc
// sites.json 里
{ "id": "03", "name": "中国裁判文书网",
  "stripPrefixes": ["发行人", "重要子公司", "担保人", "董事", "监事", "高管"] }
```

这样 `发行人：示例集团有限公司` 复制出去就是 `示例集团有限公司`。

## 从现有 Word 底稿批量生成清单

如果你已经有上一期的 Word 底稿，不用手敲。思路是把 docx 解压出 `word/document.xml`，
按段落提取标题文字，组装成上面的 JSON：

```powershell
# docx 本质是个 zip，把段落文字读出来
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead("上期底稿.docx")
$entry = $zip.Entries | Where-Object Name -eq "document.xml"
$xml = [xml](New-Object System.IO.StreamReader($entry.Open())).ReadToEnd()
$xml.document.body.p | ForEach-Object { -join ($_.r.t.'#text') } | Where-Object { $_ }
$zip.Dispose()
```

段落文字直接当 `text`，剥掉前缀当 `searchName`，加 `.jpg` 当 `imgName`，就是一份现成清单。
上一期的清单大多能直接复用，每期只改动几个人名。

---

# 配置二：增删核查网站

## 仓库自带三份清单

| 文件 | 内容 | 适合 |
|---|---|---|
| [`config/sites.example.json`](config/sites.example.json) | 4 个站 | 第一次试跑 |
| [`config/sites.gov33.json`](config/sites.gov33.json) | **33 个政府 / 司法 / 公开信息网站** | 首次尽调的全面核查 |
| [`config/sites.bond6.json`](config/sites.bond6.json) | 6 个常用站 | 债券存续期月度核查 |

用哪份：

```powershell
.\run.ps1 -Sites config\sites.gov33.json -Companies 我的清单.txt
```

## 加一个网站

用记事本或 VS Code 打开对应的 json，在数组里加一段：

```jsonc
{
  "id": "34",
  "name": "某某省市场监督管理局",
  "url": "https://scjgj.example.gov.cn/",
  "manual": true,
  "note": "站内检索入口在页面右上角，搜完点进详情页再截图"
}
```

保存即可，**不用改任何代码**。

### 字段全表

| 字段 | 必填 | 作用 |
|---|:---:|---|
| `id` | 是 | **两位数字**，决定执行顺序，也是 `-Site` / `-FromSite` / `-ToSite` 的筛选依据 |
| `name` | 是 | 显示名；纯文本清单模式下，文件夹名就是 `id-name` |
| `url` | 是 | 脚本自动打开的网址。填**搜索页**比填首页省事 |
| `manual` | 否 | 仅作提示，标记这站通常要登录或验证码，跑到时会提醒你 |
| `note` | 否 | 跑到这站时打印的提示语。**建议写清楚「搜完要点哪里再截图」**，换人接手不用问 |
| `stripPrefixes` | 否 | 数组。复制搜索词时要剥掉的前缀，见上一节 |
| `searchByCompanyOnly` | 否 | `true` 时，一家公司只在进入时复制一次搜索词 |

### `searchByCompanyOnly` 什么时候用

有些站的操作模式是「搜一次公司，然后在同一个公司页面里点不同标签页，各截一张图」——
比如先截**股权穿透图**再截**工商变更记录**。这时两张图共用一个搜索词，加上这个字段，
脚本就不会在两张图之间重复复制。

```jsonc
{ "id": "06", "name": "股权穿透核查", "url": "https://...",
  "searchByCompanyOnly": true,
  "note": "第1张=股权穿透图，第2张=工商变更记录" }
```

### 编号规则和排序

- `id` 按**字符串**比较，所以必须两位数：写 `"04"` 而不是 `"4"`，否则 `"10"` 会排在 `"4"` 前面
- 超过 99 个站的话改成三位 `"001"` 也行，但全表要统一位数
- **删网站直接删那一段就行，不用重编号** —— 脚本只按 `id` 排序，中间断号完全没问题

### 加完先单独试跑这一站

```powershell
.\run.ps1 -Sites config\sites.gov33.json -Companies 我的清单.txt -Site 34
```

确认网址能打开、搜索词复制正常、截图落到对的文件夹，再纳入整批跑。

---

# 配置三：自定义文件夹结构

## 默认长什么样

```
截图输出_20260721/              ← -Out 决定
└─ 发行人/                       ← group 决定（不填就没这层）
    └─ 示例集团有限公司/          ← name 决定
        └─ 03-裁判文书/           ← folderName 决定
            ├─ 发行人：示例集团有限公司.jpg   ← imgName 决定
            ├─ 董事长：张三.jpg
            ├─ 03-裁判文书.md                ← 自动生成，与所在文件夹同名
            └─ 03-裁判文书.docx              ← Pandoc 导出
```

## 四个控制点

| 想改什么 | 改哪里 |
|---|---|
| 最外层输出目录 | 命令行 `-Out D:\核查\2026Q3` |
| 要不要按发行人/担保人分组 | 结构化清单里的 `group` 字段，**留空或删掉就不建这层** |
| 公司文件夹名 | 结构化清单里的 `name` 字段 |
| 网站文件夹名 | 结构化清单里的 `folderName`；纯文本清单模式下自动取 `id-name` |

## 几种常见结构怎么配

**① 扁平，不分组**（适合客户背景核查）—— 清单里不写 `group`：

```
输出/
├─ 某某集团有限公司/
│   ├─ 01-信用中国/
│   └─ 02-裁判文书/
└─ 某某控股股份有限公司/
```

**② 按主体类型分组**（适合债券尽调）—— 清单里写 `"group": "发行人"` / `"担保人"`：

```
输出/
├─ 发行人/
│   └─ 某某集团有限公司/
└─ 担保人/
    └─ 某某融资担保有限公司/
```

配合 `-Group 担保人` 可以只跑其中一组。

**③ 网站文件夹不要编号前缀** —— 把 `folderName` 从 `03-裁判文书` 改成 `裁判文书` 即可。
注意 `.md` 和 `.docx` 会跟着改名，因为它们始终与所在文件夹同名。

**④ 按核查期归档** —— 每期换一个 `-Out`：

```powershell
.\run.ps1 -Out D:\核查底稿\2026-07 -Companies 我的清单.json
.\run.ps1 -Out D:\核查底稿\2026-08 -Companies 我的清单.json
```

断点续跑只看当期输出目录，所以上一期的图不会让这一期误跳过。

## 想彻底改成别的结构

目录是在 `run.ps1` 主循环里拼的，就这几行：

```powershell
$cDir = if ($c.group) { Join-Path (Join-Path $Out (Get-SafeName $c.group)) (Get-SafeName $c.name) }
        else          { Join-Path $Out (Get-SafeName $c.name) }
$sDir = Join-Path $cDir (Get-SafeName $sd.folderName)
```

想改成「网站在外、公司在内」（先按网站分文件夹，每个网站下放所有公司），换成：

```powershell
$cDir = Join-Path $Out (Get-SafeName $sd.folderName)
$sDir = Join-Path $cDir (Get-SafeName $c.name)
```

Markdown 和 Word 会跟着落到新位置，其余逻辑一行都不用动。

---

# 命令参考

```powershell
# —— 指定输入输出 ——
.\run.ps1 -Sites config\sites.gov33.json      # 换网站清单
.\run.ps1 -Companies 我的清单.json             # 换公司清单（.txt 或 .json）
.\run.ps1 -Out D:\核查\2026Q3                  # 换输出目录

# —— 筛选范围 ——
.\run.ps1 -Site 04                            # 只跑 04 号站
.\run.ps1 -Site 裁判文书                       # 站名模糊匹配
.\run.ps1 -FromSite 10 -ToSite 20             # 跑 10~20 号站
.\run.ps1 -Company 某某集团                    # 只跑某家公司（模糊匹配）
.\run.ps1 -Group 担保人                        # 只跑某一组

# —— 其他 ——
.\run.ps1 -Overwrite                          # 强制重截，覆盖已有截图
.\run.ps1 -Browser chrome                     # 用 Chrome（默认 Edge）
.\run.ps1 -NewWindow                          # 每次在新窗口打开网址
.\run.ps1 -Python "C:\Python313\python.exe"   # 指定 Python

# —— Word 导出 ——
.\export_word.ps1 -Root .\截图输出_20260721            # 补导整个目录
.\export_word.ps1 -Root .\截图输出_20260721 -Force     # 强制覆盖已有 docx
```

参数可以自由组合，例如只重跑担保人的第 4 站并覆盖旧图：

```powershell
.\run.ps1 -Group 担保人 -Site 04 -Overwrite
```

**建议按网站分批跑**，别一次跑几百张。先跑不用登录的站（比如百度）找找手感，再跑需要登录的。
每跑完一站抽查几张图确认没问题，再跑下一站。

---

# 常见问题

**跑到一半浏览器崩了 / 登录态掉了**
按 `q` 退出，重新登录，再跑同样的命令。已截的图自动跳过，从断点继续。

**某张图截错了**
当场按 `r` 重截。已经跑过去了的话，删掉那张 jpg 再跑会自动补上，或者用 `-Overwrite` 重跑这家。

**某个网站搜不到结果**
搜不到也要截图存证（证明查过、无记录）。直接按回车截当前的「无结果」页面即可。

**截出来的图里有另一块屏或桌面**
确认用的是本仓库的 `capture_screen.py`，它只截浏览器所在的那块屏。仍有问题的话，
用 `-Python` 确认调用的是装了 Pillow 的那个 Python。

**截出来的字很糊**
把浏览器放到高 DPI 屏上截，或在网页上按 `Ctrl` `+` 放大到 125% 再截。

**改了 json 但没生效**
脚本在**启动时**读一次配置，改完要重新启动才生效。另外确认 json 存为 **UTF-8** 编码，
否则中文会乱码。

**想确认 json 没写错**
```powershell
Get-Content config\sites.gov33.json -Raw -Encoding UTF8 | ConvertFrom-Json
```
不报错就是合法的。

---

## 合规提示

- 本工具只做**公开信息**的检索留痕，请在各网站的服务条款和使用规则允许的范围内使用
- 不要用它做批量数据抓取；它一次只处理一个页面，且每一步都需要人工确认
- 验证码请自行完成，不要接入打码平台 —— 那既违反网站条款，也会让核查底稿的可信度存疑

## License

MIT，见 [LICENSE](LICENSE)。
