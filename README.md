# Rename-Subtitles / 字幕批量对齐脚本

一个用于**批量把字幕文件名对齐到对应视频**的 PowerShell 脚本。

适合这类场景：

- 视频和字幕命名风格不一致
- 文件名里混有 `1080p`、`HEVC`、`10bit`、`WEB-DL`、`SCJP` 等发布信息
- 简中 / 繁中 / 日语 / 英语字幕混在一起
- 字幕放在 `sc`、`tc`、`jp` 等子目录里

> 脚本**只处理字幕文件**，不会修改视频文件。

---

## 核心规则

1. **同目录默认视为同一季**，主要按“集数”匹配。
2. `-VideoQuery` / `-SubtitleQuery` 只做**粗筛**，最终仍靠集数和文件名相似度判断。
3. 语种优先级固定为：

   > **文件名 > 所在语言目录 > `-LanguageHint`**

4. 遇到不确定情况（如找不到集数、候选视频太多、目标文件已存在）会**跳过**，不会乱改。
5. 第一次使用建议始终先加 `-WhatIf` 预览。

---

## 内置识别能力

### 集数格式

默认支持：

- `S01E01`
- `E01`
- `[01]`
- `第01集` / `第1話`
- `01` / `1`
- 全角数字：`１` / `２`

### 常见语种写法

| 最终输出 | 常见别名示例 |
|---|---|
| `JP` | `jp`、`jpn`、`ja`、`japanese`、`日本語`、`日语`、`日文` |
| `EN` | `en`、`eng`、`english`、`英文`、`英语` |
| `SC` | `sc`、`chs`、`zhs`、`zh-cn`、`zh-hans`、`gb`、`简中`、`简体中文` |
| `TC` | `tc`、`cht`、`zht`、`zh-tw`、`zh-hant`、`big5`、`繁中`、`繁體中文` |

支持 `SCJP`、`BIG5JP`、`name.jp.ass`、`name_zh-cn.ass` 这类常见风格。

### 语种输出规则

- 全大写
- 不加分隔符
- 中文放最后

例如：`SCJP -> JPSC`、`JP + TC -> JPTC`

### 自动忽略的常见噪声

会自动清理这类发布信息，避免干扰集数判断：

- 分辨率：`480p` `720p` `1080p` `2160p` `4K`
- 片源：`WEB-DL` `WEBRip` `BluRay` `BDRip` `Remux`
- 编码：`x264` `x265` `H264` `H265` `HEVC` `AV1`
- 其他：`10bit` `HDR` `DV` `DoVi` `AAC` `DTS` `Atmos` `SDH` `CC`

---

## 快速开始

### 1. 先预览

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -WhatIf
```

### 2. 确认无误后正式执行

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime"
```

### 3. 字幕在子目录里时递归处理

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -Recurse -WhatIf
```

---

## 使用方法

### 基本格式

```powershell
.\Rename-Subtitles.ps1 `
  [-Directory <路径>] `
  [-VideoQuery <视频粗筛关键词>] `
  [-SubtitleQuery <字幕粗筛关键词>] `
  [-LanguageHint <语种提示>] `
  [-VideoEpisodePatterns <正则数组>] `
  [-SubtitleEpisodePatterns <正则数组>] `
  [-Recurse] `
  [-WhatIf]
```

一行写法也可以：

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -WhatIf
```

---

## 可选：封装成 `pwsh` 里的命令

如果你经常使用，建议把它写进 PowerShell Profile。这样打开 `pwsh` 后可以直接用 `rename-subtitles` 或 `rs`。

### 查看并创建 `$PROFILE`

```powershell
$PROFILE

if (-not (Test-Path $PROFILE)) {
    New-Item -ItemType File -Path $PROFILE -Force | Out-Null
}
```

### 推荐封装方式（保留参数补全）

```powershell
function rename-subtitles {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$Directory = ".",
        [string]$VideoQuery = "",
        [string]$SubtitleQuery = "",
        [string]$LanguageHint = "",
        [string[]]$VideoEpisodePatterns = @(),
        [string[]]$SubtitleEpisodePatterns = @(),
        [switch]$Recurse
    )

    $params = @{
        Directory               = $Directory
        VideoQuery              = $VideoQuery
        SubtitleQuery           = $SubtitleQuery
        LanguageHint            = $LanguageHint
        VideoEpisodePatterns    = $VideoEpisodePatterns
        SubtitleEpisodePatterns = $SubtitleEpisodePatterns
        Recurse                 = $Recurse
    }

    if ($WhatIfPreference) {
        $params.WhatIf = $true
    }

    & "E:\Study\Powershell\RenameSubtitle\Rename-Subtitles.ps1" @params
}

Set-Alias rs rename-subtitles
```

保存后执行：

```powershell
. $PROFILE
```

以后就可以直接这样用：

```powershell
rename-subtitles -Directory "D:\Anime" -WhatIf
rs -Directory "D:\Anime" -WhatIf
```

> 不建议用 `ArgsList` 全量透传的简写函数；那种写法通常拿不到底层参数的自动补全。

---

## 参数说明

| 参数 | 作用 |
|---|---|
| `-Directory` | 指定处理目录，默认是当前目录 |
| `-VideoQuery` | 粗筛视频文件，缩小候选范围 |
| `-SubtitleQuery` | 粗筛字幕文件，只筛文件，不负责语种判断 |
| `-LanguageHint` | 当文件名和目录都识别不出语种时，提供兜底语种 |
| `-VideoEpisodePatterns` | 自定义视频文件名的集数正则 |
| `-SubtitleEpisodePatterns` | 自定义字幕文件名的集数正则 |
| `-Recurse` | 递归扫描子目录 |
| `-WhatIf` | 预览模式，不真正修改文件 |

### 常见参数示例

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -SubtitleQuery "big5" -WhatIf
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -LanguageHint "zh-cn" -WhatIf
.\Rename-Subtitles.ps1 -Directory "D:\Mixed" -VideoQuery "Hotel" -SubtitleQuery "Hotel" -WhatIf
```

---

## 手动传入正则的示例

只有默认规则识别不好时，才建议手动传正则。

### 案例 1：视频是 `Episode-01`，字幕是 `第01话`

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" `
  -VideoEpisodePatterns @(
    '(?i)Episode[-_ ](?<episode>\d{1,2})'
  ) `
  -SubtitleEpisodePatterns @(
    '(?i)第\s*(?<episode>\d{1,2})\s*[话話]'
  ) `
  -WhatIf
```

### 案例 2：视频是 `EP01`，字幕是 `01v2`

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" `
  -VideoEpisodePatterns @(
    '(?i)\bEP(?<episode>\d{1,2})\b'
  ) `
  -SubtitleEpisodePatterns @(
    '(?i)(?<episode>\d{1,2})v\d+'
  ) `
  -WhatIf
```

### 使用建议

1. 尽量使用命名捕获组：`(?<episode>\d{1,2})`
2. 多套规则时，把**最准确的放前面**
3. 一定先配合 `-WhatIf` 预览

---

## 输出状态说明

| 状态 | 含义 |
|---|---|
| `Renamed` | 已在原目录完成重命名 |
| `MovedRenamed` | 已从子目录移动到视频目录并重命名 |
| `Unchanged` | 文件名本来就已经是目标格式，无需处理 |
| `SkippedExists` | 目标文件已存在，已跳过 |
| `Ambiguous` | 同一集有多个候选视频，脚本无法安全判断 |
| `Error` | 处理该字幕时发生异常 |
| `WhatIf` | 预览模式，仅展示结果，未真正修改 |

---

## 注意事项

1. 第一次使用建议先加 `-WhatIf`
2. 默认按“同目录同一季”处理，不单独判断季度
3. 视频文件不会被改动
4. `-LanguageHint` 是兜底提示，不是强制覆盖
5. 如果完全识别不出集数，脚本就无法自动配对
