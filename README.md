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

1. 视频和字幕先按扩展名及 `-VideoQuery` / `-SubtitleQuery` 筛选，再分别提取集数并按整数匹配。
2. `-VideoQuery` / `-SubtitleQuery` 是文件名包含关键词的筛选条件；文件名相似度只在同一集有多个候选视频时用于消歧。
3. 语种优先级固定为：

   > **文件名 > 所在语言目录 > `-LanguageHint`**

4. 同一集的多语种字幕在自动偏移计数时只算一个集数。
5. 遇到无法识别集数、没有匹配视频、候选视频无法消歧或目标文件已存在时，不会强行重命名。
6. `-Recurse` 会将扫描到的文件作为一个处理批次；脚本不解析季号。跨季度且集数重置时，请按季度分别运行。
7. 第一次使用建议先加 `-WhatIf` 预览。

---

## 内置识别能力

### 集数格式

默认支持：

- `S01E01`
- `E01`
- `[01]`
- `第01集` / `第1話`
- `01` / `1`
- `001` / `0001` 等带任意位数前导零的集数（先去掉前导零，再按正整数匹配；数值需在 `Int32` 范围内）
- 全角数字：`１` / `２`

无分隔符的裸数字需要位于文件名开头，或由空格、点、下划线、连字符、方括号、圆括号等分隔；例如 `abc 001.ass` 可识别。自定义正则中的集数捕获也应使用命名组 `(?<episode>...)`，集数文本需能转换为 `Int32`。

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
  [-EpisodeAdd <手动字幕集数偏移 int>] `
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
        [int]$EpisodeAdd = 0,
        [switch]$Recurse
    )

    $params = @{
        Directory               = $Directory
        VideoQuery              = $VideoQuery
        SubtitleQuery           = $SubtitleQuery
        LanguageHint            = $LanguageHint
        VideoEpisodePatterns    = $VideoEpisodePatterns
        SubtitleEpisodePatterns = $SubtitleEpisodePatterns
        EpisodeAdd              = $EpisodeAdd
        Recurse                 = $Recurse
    }

    if ($WhatIfPreference) {
        $params.WhatIf = $true
    }

    & "D:\Tools\Rename-Subtitles.ps1" @params
}

Set-Alias rs rename-subtitles
```

将示例中的 `D:\Tools\Rename-Subtitles.ps1` 替换为脚本在本机的实际路径。

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
| `-EpisodeAdd` | 字幕集数偏移。默认 `0` 时尝试自动推导；非零值覆盖自动推导。正数增加、负数减少 |
| `-Recurse` | 递归扫描子目录 |
| `-WhatIf` | 预览模式，不真正修改文件 |

### 常见参数示例

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -SubtitleQuery "big5" -WhatIf
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -LanguageHint "zh-cn" -WhatIf
.\Rename-Subtitles.ps1 -Directory "D:\Mixed" -VideoQuery "Hotel" -SubtitleQuery "Hotel" -WhatIf
```

### 自动推导集数偏移

默认情况下（`-EpisodeAdd 0`），脚本会比较筛选到的视频和字幕中的**不同集数**。两边集数数量相同、所有候选文件都能识别集数，且分别排序后逐项计算出的差值一致时，自动使用该差值匹配。相同集数的多语种字幕只计一次。

偏移的计算方向是“视频集数 - 字幕集数”。例如视频集数为 `1..12`、字幕集数为 `13..24`，脚本会推导偏移 `-12`，再将偏移加到每个字幕集数上进行视频匹配。若数量不一致、存在无法识别的候选文件或逐项差值不一致，则不自动偏移，按零偏移继续尝试匹配；如需强制偏移，可手动指定：

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime\Season02" -EpisodeAdd -12 -WhatIf
```

非零 `-EpisodeAdd` 会直接覆盖自动推导。自动推导只针对本次筛选到的文件批次；若递归扫描多个季度且每季集数都从 1 重新开始，脚本不会按目录名识别季别，还可能出现重复集数的多个候选视频。建议按季度分别运行。

结果中的集数键会按本次识别到的最长集数字符串补零。实际字幕文件名以匹配视频的主体名为准，识别到语种时再追加语种标签；处理语言子目录中的字幕时，字幕可能会移动到视频所在目录。

---

## 手动传入正则的示例

只有默认规则识别不好时，才建议手动传正则。

### 案例 1：视频是 `Episode-01`，字幕是 `第01话`

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" `
  -VideoEpisodePatterns @(
    '(?i)Episode[-_ ](?<episode>\d+)'
  ) `
  -SubtitleEpisodePatterns @(
    '(?i)第\s*(?<episode>\d+)\s*[话話]'
  ) `
  -WhatIf
```

### 案例 2：视频是 `EP01`，字幕是 `01v2`

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" `
  -VideoEpisodePatterns @(
    '(?i)\bEP(?<episode>\d+)\b'
  ) `
  -SubtitleEpisodePatterns @(
    '(?i)(?<episode>\d+)v\d+'
  ) `
  -WhatIf
```

### 使用建议

1. 尽量使用命名捕获组：`(?<episode>\d+)`
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
