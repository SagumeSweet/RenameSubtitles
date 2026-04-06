# Rename Subtitle Scripts

一个用于 **对齐字幕** 的 PowerShell 脚本。

## 功能概览

### 1. 按集数匹配视频与字幕

脚本默认认为：

> **同一个目录中的视频属于同一季，因此只按“集数”匹配，不处理季度问题。**

支持识别的集数形式包括：

- `S01E01`
- `E01`
- `[01]`
- `第01集`
- `第1話`
- `1` / `2`
- 全角数字：`１` / `２`

识别后会统一转换成内部键值，例如：

- `E01`
- `E02`

---

### 2. 用 PowerShell 通配符先做粗筛

脚本会先根据你传入的关键词，结合扩展名白名单生成通配符数组，然后再用正则做精确匹配。

例如：

- `-VideoQuery "1080p"`
- `-SubtitleQuery "big5"`

这类参数的作用是：

- **缩小候选文件范围**
- 不是最终判断依据

---

### 3. 自动识别语种标签

脚本内置语种别名表，支持从字幕文件名中识别：

| 输出标签 | 可识别别名示例 |
|---|---|
| `JP` | `jp`, `jpn`, `ja`, `japanese`, `日语`, `日文` |
| `SC` | `sc`, `chs`, `gb`, `cn`, `简中`, `简体`, `简体中文` |
| `TC` | `tc`, `cht`, `big5`, `繁中`, `繁體`, `繁體中文` |

支持的命名风格包括：

- `name.jp.ass`
- `name-jp.ass`
- `name_jp.ass`
- `[big5]`
- `SCJP`
- `CHSJP`
- `BIG5JP`

### 输出规则

- 无分隔符
- 统一大写
- 中文标签放最后

例如：

- `SCJP` -> `JPSC`
- `JP + TC` -> `JPTC`

---

### 4. 自动清洗视频发布信息噪声

为了避免把 `1080p`、`x265`、`10bit` 等内容误识别成集数，脚本会自动清洗常见发布参数，例如：

- `1080p`, `720p`, `2160p`
- `x264`, `x265`, `h264`, `h265`, `hevc`, `avc`
- `10bit`, `8bit`
- `WEB-DL`, `WEBRip`, `BDRip`, `Bluray`, `Remux`
- `AAC`, `FLAC`, `DTS`, `TrueHD`, `Atmos`

---

### 5. 支持语言子文件夹自动处理

如果目录中存在类似下面的结构：

```text
Anime 01.mkv
Anime 02.mkv
sc\[SubGroup] Anime [01].ass
tc\[SubGroup] Anime [02].ass
```

脚本运行时会检测到：

- `sc`
- `tc`
- `jp`

这类语言子目录，并在运行中提示：

> 是否自动搜索这些子目录中的字幕，并移动到视频所在目录旁边。

如果你加上 `-Recurse`，则会直接递归处理，不再提示。

---

### 6. 不会修改视频文件

脚本**不会重命名、不会移动、不会修改任何视频文件**。

视频文件仅用于：

- 识别集数
- 提供目标文件名模板

真正被改动的只有字幕文件。

---

## 文件说明

当前项目包含：

- `Rename-Subtitles.ps1`：主脚本
- `Install-RenameSubtitles-Command.ps1`：用于把主脚本封装成 PowerShell 命令

---

## 安装为命令

如果你想把脚本封装成命令，在 PowerShell 中执行：

```powershell
& "I:\ProgramFiles\PowershellScript\RenameSubtitle\Install-RenameSubtitles-Command.ps1"
```

执行后，你就可以直接使用：

```powershell
rename-subtitles -Directory . -WhatIf
```

或者使用简写：

```powershell
rs -Directory . -WhatIf
```

---

## 详细用法

### 基本格式

```powershell
.\Rename-Subtitles.ps1 \
  [-Directory <路径>] \
  [-VideoQuery <视频粗筛关键词>] \
  [-SubtitleQuery <字幕粗筛关键词>] \
  [-LanguageHint <语种提示>] \
  [-VideoEpisodePatterns <正则数组>] \
  [-SubtitleEpisodePatterns <正则数组>] \
  [-Recurse] \
  [-WhatIf]
```

---

## 参数说明

### `-Directory`

要处理的目录，默认是当前目录。

示例：

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime"
```

---

### `-VideoQuery`

用于粗筛视频文件的关键词。

例如只处理名称里带 `1080p` 的视频：

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -VideoQuery "1080p"
```

---

### `-SubtitleQuery`

用于粗筛字幕文件的关键词。

例如只处理带 `big5` 的字幕：

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -SubtitleQuery "big5"
```

> 注意：它现在只用于筛字幕文件，不再兼任语种提示。

---

### `-LanguageHint`

当字幕文件名和目录都无法识别语种时，手动指定一个语种提示。

例如：

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -LanguageHint "big5"
```

那么未识别语种的字幕，会优先尝试按 `TC` 处理。

---

### `-VideoEpisodePatterns`

自定义视频文件名的集数提取正则数组。

一般不需要改，除非你的视频命名特别特殊。

---

### `-SubtitleEpisodePatterns`

自定义字幕文件名的集数提取正则数组。

例如：

```powershell
.\Rename-Subtitles.ps1 -SubtitleEpisodePatterns @(
  '(?i)\[(?<episode>\d{1,2})\]',
  '(?i)第\s*(?<episode>\d{1,2})\s*[集话話]'
)
```

---

### `-Recurse`

递归扫描子目录。

当字幕被解压在 `sc` / `tc` / `jp` 等子目录里时，推荐使用。

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -Recurse -WhatIf
```

---

### `-WhatIf`

预览模式，不会真正修改文件。

这是第一次使用时**强烈推荐**的参数：

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -WhatIf
```

---

## 常见示例

### 示例 1：预览当前目录中的重命名结果

```powershell
.\Rename-Subtitles.ps1 -WhatIf
```

---

### 示例 2：处理指定目录

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime"
```

---

### 示例 3：只处理 `big5` 字幕

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -SubtitleQuery "big5" -WhatIf
```

---

### 示例 4：递归处理语言子文件夹

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -Recurse -WhatIf
```

---

### 示例 5：手动指定语种提示

```powershell
.\Rename-Subtitles.ps1 -Directory "D:\Anime" -LanguageHint "jp" -WhatIf
```

---

## 输出状态说明

脚本执行后会输出结果对象，并附带汇总统计。

常见状态如下：

| 状态 | 含义 |
|---|---|
| `Renamed` | 已在原目录中完成重命名 |
| `MovedRenamed` | 已从子目录移动到视频目录并完成重命名 |
| `Unchanged` | 新旧文件名相同，无需处理 |
| `SkippedExists` | 目标文件已存在，已跳过 |
| `Ambiguous` | 有多个候选视频，无法自动判断 |
| `Error` | 处理单个字幕时出现异常 |
| `WhatIf` | 预览模式，未真正改动文件 |

---

## 结果汇总示例

```text
处理结果汇总：
  Renamed       : 5
  MovedRenamed  : 2
  SkippedExists : 1
  Ambiguous     : 1
  Error         : 0
```

---

## 注意事项

1. **第一次使用建议加 `-WhatIf`**
2. 默认按“同目录同一季”处理，不考虑季度匹配
3. 视频文件不会被修改
4. 如果某些字幕完全没有集数信息，脚本无法自动匹配
5. 如果同一集存在多个视频候选，脚本会尽量用相似度判断；无法区分时会跳过并提示

---

## 建议的使用顺序

### 第一步：预览

```powershell
rename-subtitles -Directory "你的目录" -WhatIf
```

### 第二步：确认输出没问题后正式执行

```powershell
rename-subtitles -Directory "你的目录"
```

### 第三步：遇到子目录字幕时使用递归

```powershell
rename-subtitles -Directory "你的目录" -Recurse -WhatIf
```
