<#
.SYNOPSIS
    按视频文件名批量整理同目录字幕文件。

.DESCRIPTION
    该脚本会在指定目录中查找视频与字幕，按“集数”进行匹配，
    再将字幕重命名为与对应视频相同的主体名称，并在需要时追加语种标签。

    设计目标：
    1. 只处理字幕，不修改任何视频文件。
    2. 同一个目录默认视为同一季，只按集数匹配。
    3. 先用 PowerShell 通配符粗筛，再用正则精确提取集数与语种。
    4. 如果检测到 `sc` / `tc` / `jp` 等语言子文件夹，会在运行时询问是否自动处理。

.PARAMETER Directory
    要处理的目录，默认是当前目录。

.PARAMETER VideoQuery
    视频文件粗筛关键字，会与扩展名一起生成通配符。

.PARAMETER SubtitleQuery
    字幕文件粗筛关键字，仅用于筛选候选字幕文件。

.PARAMETER LanguageHint
    可选的语种提示。当字幕文件名和所在目录都无法识别语种时，使用该值作为回退标签。

.PARAMETER VideoEpisodePatterns
    用于从视频文件名提取集数的正则表达式数组。

.PARAMETER SubtitleEpisodePatterns
    用于从字幕文件名提取集数的正则表达式数组。

.PARAMETER Recurse
    是否递归扫描子目录。

.EXAMPLE
    .\Rename-Subtitles.ps1 -Directory 'D:\Anime' -WhatIf

.EXAMPLE
    .\Rename-Subtitles.ps1 -Directory 'D:\Anime' -SubtitleQuery 'big5'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [string]$Directory = ".",
    [string]$VideoQuery = "",
    [string]$SubtitleQuery = "",
    [string]$LanguageHint = "",
    [string[]]$VideoEpisodePatterns = @(
        '(?i)\bS(?<season>\d{1,2})[ ._-]*E(?<episode>\d{1,3})\b',
        '(?i)\bE(?<episode>\d{1,3})\b',
        '(?i)\[(?<episode>\d{1,3})\]',
        '(?i)第\s*(?<episode>\d{1,3})\s*[集话話]',
        '(?i)(?:^|[\s._\-\[\(])(?<episode>\d{1,2})(?:$|[\s._\-\]\)])'
    ),
    [string[]]$SubtitleEpisodePatterns = @(
        '(?i)\bS(?<season>\d{1,2})[ ._-]*E(?<episode>\d{1,3})\b',
        '(?i)\bE(?<episode>\d{1,3})\b',
        '(?i)\[(?<episode>\d{1,3})\]',
        '(?i)第\s*(?<episode>\d{1,3})\s*[集话話]',
        '(?i)(?:^|[\s._\-\[\(])(?<episode>\d{1,2})(?:$|[\s._\-\]\)])'
    ),
    [switch]$Recurse
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region 固定配置
# 常见视频扩展名白名单
[string[]]$script:VideoExtensions = @(
    'mp4', 'mkv', 'avi', 'mov', 'flv', 'wmv', 'mpeg', 'mpg', 'webm', 'vob', 'ogv', 'ogg', '3gp', '3g2', 'm4v'
)

# 常见字幕扩展名白名单
[string[]]$script:SubtitleExtensions = @(
    'srt', 'ass', 'ssa', 'sub', 'idx', 'vtt', 'ttml', 'dfxp', 'smi', 'sami', 'aqt', 'mpl2', 'pjs', 'stl', 'usf', 'jss', 'sbv', 'lrc'
)

# 语种别名表：值越全，识别越稳
$script:LanguageRules = [ordered]@{
    jp = @('jp', 'jpn', 'ja', 'japanese', '日语', '日文')
    sc = @('sc', 'chs', 'gb', 'cn', '简中', '简体', '简体中文')
    tc = @('tc', 'cht', 'big5', '繁中', '繁體', '繁體中文')
}
#endregion

function New-SearchPatterns {
    <#
        根据用户输入的关键词和扩展名，拼出 `Get-ChildItem -Include` 可用的通配符数组。
    #>
    param(
        [string]$Query,
        [string[]]$Extensions
    )

    return @($Extensions | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace($Query)) {
            "*.$_"
        }
        else {
            "*$Query*.$_"
        }
    })
}

function Convert-ToHalfWidthDigits {
    <#
        将全角数字（如 １、２）转成半角数字（1、2），避免正则漏匹配。
    #>
    param([string]$Text)

    return [regex]::Replace($Text, '[０-９]', {
        param($match)
        [char]([int][char]$match.Value - 65248)
    })
}

function Remove-ReleaseNoise {
    <#
        去掉常见发布信息噪声，减少把 1080p / x265 / 10bit 等误识别成集数的概率。
    #>
    param([string]$Text)

    return [regex]::Replace(
        $Text,
        '(?i)(?:\b\d{3,4}p\b|\b(?:x26[45]|h\.?26[45]|hevc|avc)\b|\b(?:10|8)bit\b|\b(?:aac|flac|dts|truehd|atmos)\b|\b(?:webrip|web[- ]?dl|bdrip|bluray|remux)\b|\b(?:uhd|hdr|dv)\b)',
        ' '
    )
}

function Get-EpisodeKey {
    <#
        从文件名中提取集数，并统一格式为 E01 / E02。
        当前脚本按“同目录同一季”设计，因此只保留集数，不保留季号。
    #>
    param(
        [string]$Name,
        [string[]]$Patterns
    )

    $normalizedName = Convert-ToHalfWidthDigits -Text $Name
    $candidateName = Remove-ReleaseNoise -Text $normalizedName

    foreach ($pattern in $Patterns) {
        $match = [regex]::Match($candidateName, $pattern)
        if (-not $match.Success) {
            continue
        }

        $episodeGroup = $match.Groups['episode']
        if ($episodeGroup.Success) {
            return ('E{0:D2}' -f [int]$episodeGroup.Value)
        }

        $captureValues = @()
        for ($index = 1; $index -lt $match.Groups.Count; $index++) {
            if ($match.Groups[$index].Success -and $match.Groups[$index].Value -match '^\d+$') {
                $captureValues += $match.Groups[$index].Value
            }
        }

        if ($captureValues.Count -ge 1) {
            return ('E{0:D2}' -f [int]$captureValues[0])
        }
    }

    return $null
}

function Get-LanguageTag {
    <#
        从文件名中识别语种标签。
        支持：
        - 分隔形式：.jp / -tc / _chs / [big5]
        - 紧凑组合：SCJP / CHSJP / BIG5JP

        输出规则：
        - 无分隔符
        - 统一大写
        - 中文标签放最后，如 JPSC / JPTC
    #>
    param(
        [string]$Name,
        [System.Collections.IDictionary]$Rules
    )

    $detectedTags = New-Object System.Collections.Generic.List[string]
    $aliasEntries = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $Rules.GetEnumerator()) {
        foreach ($alias in $entry.Value) {
            $aliasEntries.Add([PSCustomObject]@{
                Canonical = [string]$entry.Key
                Alias     = $alias.ToLowerInvariant()
                Length    = $alias.Length
            }) | Out-Null
        }
    }

    $sortedAliasEntries = $aliasEntries | Sort-Object Length -Descending

    foreach ($entry in $Rules.GetEnumerator()) {
        foreach ($alias in $entry.Value) {
            $escapedAlias = [regex]::Escape($alias)
            $boundaryPattern = "(?i)(?:^|[\.\-_\s\[\]\(\)\{\}])$escapedAlias(?:$|[\.\-_\s\[\]\(\)\{\}])"
            if ($Name -match $boundaryPattern -and -not $detectedTags.Contains([string]$entry.Key)) {
                $detectedTags.Add([string]$entry.Key) | Out-Null
            }
        }
    }

    $normalizedName = $Name.ToLowerInvariant()
    $tokens = [regex]::Split($normalizedName, '[\.\-_\s\[\]\(\)\{\}]+') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    foreach ($token in $tokens) {
        $rest = $token
        while ($rest.Length -gt 0) {
            $matchedAlias = $sortedAliasEntries | Where-Object { $rest.StartsWith($_.Alias) } | Select-Object -First 1
            if (-not $matchedAlias) {
                break
            }

            if (-not $detectedTags.Contains($matchedAlias.Canonical)) {
                $detectedTags.Add($matchedAlias.Canonical) | Out-Null
            }
            $rest = $rest.Substring($matchedAlias.Alias.Length)
        }
    }

    if ($detectedTags.Count -eq 0) {
        return $null
    }

    $priority = @('jp', 'sc', 'tc')
    $sortedTags = $detectedTags | Sort-Object { $priority.IndexOf($_) }
    $nonChineseTags = @($sortedTags | Where-Object { $_ -notin @('sc', 'tc') })
    $chineseTags = @($sortedTags | Where-Object { $_ -in @('sc', 'tc') })
    $finalTags = @($nonChineseTags + $chineseTags)

    return (($finalTags -join '').ToUpperInvariant())
}

function Get-ExactLanguageTag {
    <#
        用于“语言子目录自动发现”的严格匹配函数。
        只有当整个目录名都能被语种别名完整解释时，才认定它是语言目录。

        例如：
        - `sc` -> `SC`
        - `tc` -> `TC`
        - `scjp` -> `JPSC`

        而像下面这些不会被识别为语言目录：
        - `subtitle-sc`
        - `backup_tc`
        - `my-jp-folder`
    #>
    param(
        [string]$Name,
        [System.Collections.IDictionary]$Rules
    )

    $normalizedName = $Name.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($normalizedName)) {
        return $null
    }

    $aliasEntries = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $Rules.GetEnumerator()) {
        $aliasEntries.Add([PSCustomObject]@{
            Canonical = [string]$entry.Key
            Alias     = ([string]$entry.Key).ToLowerInvariant()
            Length    = ([string]$entry.Key).Length
        }) | Out-Null

        foreach ($alias in $entry.Value) {
            $aliasEntries.Add([PSCustomObject]@{
                Canonical = [string]$entry.Key
                Alias     = $alias.ToLowerInvariant()
                Length    = $alias.Length
            }) | Out-Null
        }
    }

    $sortedAliasEntries = $aliasEntries | Sort-Object Length -Descending
    $detectedTags = New-Object System.Collections.Generic.List[string]
    $tokens = @([regex]::Split($normalizedName, '[\.\-_\s\[\]\(\)\{\}]+') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($tokens.Count -eq 0) {
        return $null
    }

    foreach ($token in $tokens) {
        $rest = $token
        while ($rest.Length -gt 0) {
            $matchedAlias = $sortedAliasEntries | Where-Object { $rest.StartsWith($_.Alias) } | Select-Object -First 1
            if (-not $matchedAlias) {
                return $null
            }

            if (-not $detectedTags.Contains($matchedAlias.Canonical)) {
                $detectedTags.Add($matchedAlias.Canonical) | Out-Null
            }
            $rest = $rest.Substring($matchedAlias.Alias.Length)
        }
    }

    if ($detectedTags.Count -eq 0) {
        return $null
    }

    $priority = @('jp', 'sc', 'tc')
    $sortedTags = $detectedTags | Sort-Object { $priority.IndexOf($_) }
    $nonChineseTags = @($sortedTags | Where-Object { $_ -notin @('sc', 'tc') })
    $chineseTags = @($sortedTags | Where-Object { $_ -in @('sc', 'tc') })
    $finalTags = @($nonChineseTags + $chineseTags)

    return (($finalTags -join '').ToUpperInvariant())
}

function Get-MatchTokens {
    <#
        从文件名中抽出用于“相似度匹配”的词项。
        这里会剔除：
        - 发布信息噪声
        - 已识别的语种别名
        - 已识别的集数片段
    #>
    param(
        [string]$Name,
        [System.Collections.IDictionary]$Rules
    )

    $normalizedName = Convert-ToHalfWidthDigits -Text $Name
    $cleanName = [regex]::Replace(
        (Remove-ReleaseNoise -Text $normalizedName),
        '(?i)(?:\bS\d{1,2}[ ._-]*E\d{1,3}\b|\bE\d{1,3}\b|第\s*\d{1,3}\s*[集话話])',
        ' '
    )

    $languageAliases = @()
    foreach ($entry in $Rules.GetEnumerator()) {
        $languageAliases += ([string]$entry.Key).ToLowerInvariant()
        $languageAliases += @($entry.Value | ForEach-Object { $_.ToLowerInvariant() })
    }

    [string[]]$tokens = @(
        [regex]::Split($cleanName.ToLowerInvariant(), '[^\p{L}\p{Nd}]+') | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_) -and
            $_ -notmatch '^\d+$' -and
            $_.Length -gt 1 -and
            $_ -notin $languageAliases
        } | Select-Object -Unique
    )

    return $tokens
}

function Select-BestVideoMatch {
    <#
        当同一集数对应多个视频时，根据文件名词项相似度选择最合适的视频。
        如果无法明确区分，返回 $null，由主流程标记为 Ambiguous。
    #>
    param(
        [object[]]$Videos,
        [string]$SubtitleName,
        [System.Collections.IDictionary]$Rules
    )

    [object[]]$videoArray = @($Videos)
    if ($videoArray.Count -eq 0) {
        return $null
    }

    if ($videoArray.Count -eq 1) {
        return $videoArray[0]
    }

    [string[]]$subtitleTokens = @(Get-MatchTokens -Name $SubtitleName -Rules $Rules)
    $scoredCandidates = foreach ($video in $videoArray) {
        [string[]]$videoTokens = @(Get-MatchTokens -Name $video.BaseName -Rules $Rules)
        [string[]]$sharedTokens = @($subtitleTokens | Where-Object { $videoTokens -contains $_ } | Select-Object -Unique)
        [string[]]$subtitleOnlyTokens = @($subtitleTokens | Where-Object { $videoTokens -notcontains $_ } | Select-Object -Unique)
        [string[]]$videoOnlyTokens = @($videoTokens | Where-Object { $subtitleTokens -notcontains $_ } | Select-Object -Unique)

        [PSCustomObject]@{
            Video = $video
            Score = ($sharedTokens.Count * 10) - ($videoOnlyTokens.Count * 3) - $subtitleOnlyTokens.Count
        }
    }

    $orderedCandidates = @(
        $scoredCandidates | Sort-Object -Property @(
            @{ Expression = { $_.Score }; Descending = $true },
            @{ Expression = { $_.Video.Name.Length }; Ascending = $true }
        )
    )

    if ($orderedCandidates.Count -eq 1) {
        return $orderedCandidates[0].Video
    }

    if ($orderedCandidates[0].Score -gt $orderedCandidates[1].Score) {
        return $orderedCandidates[0].Video
    }

    return $null
}

function Get-FolderLanguageTag {
    <#
        如果字幕位于语言子文件夹中（如 sc / tc / jp），则从路径继承语言标签。
    #>
    param(
        [string]$Path,
        [object[]]$LanguageFolders
    )

    foreach ($folder in $LanguageFolders | Sort-Object { $_.Path.Length } -Descending) {
        if ($Path.StartsWith($folder.Path, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $folder.Language
        }
    }

    return $null
}

function Get-LanguageFolders {
    <#
        扫描顶层目录，找出可能的语言子目录，供运行时提示使用。
        这里采用“目录名全文匹配”策略，避免因为模糊包含 `sc/tc/jp` 而误判。
    #>
    param(
        [string]$RootPath,
        [System.Collections.IDictionary]$Rules
    )

    return @(
        Get-ChildItem -LiteralPath $RootPath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $folderLanguageTag = Get-ExactLanguageTag -Name $_.Name -Rules $Rules
            if ($folderLanguageTag) {
                [PSCustomObject]@{
                    Name     = $_.Name
                    Path     = $_.FullName
                    Language = $folderLanguageTag
                }
            }
        }
    )
}

function Get-FilesSafely {
    <#
        统一封装文件枚举，出错时抛出更明确的错误信息。
    #>
    param(
        [string]$SearchPath,
        [string[]]$Patterns,
        [bool]$Recursive,
        [string]$Label
    )

    try {
        if ($Recursive) {
            return @(Get-ChildItem -Path $SearchPath -File -Recurse -Include $Patterns)
        }

        return @(Get-ChildItem -Path $SearchPath -File -Include $Patterns)
    }
    catch {
        throw "枚举$Label失败：$($_.Exception.Message)"
    }
}

function New-ResultObject {
    <#
        统一创建结果对象，避免主流程里重复拼装 PSCustomObject。
    #>
    param(
        [string]$Status,
        [string]$EpisodeKey,
        [string]$Language,
        [string]$OldName,
        [string]$NewName
    )

    return [PSCustomObject]@{
        Status     = $Status
        EpisodeKey = $EpisodeKey
        Language   = $Language
        OldName    = $OldName
        NewName    = $NewName
    }
}

function Resolve-SubtitleTarget {
    <#
        根据字幕、视频、语种标签与目录规则，生成最终目标文件名和目标路径。
    #>
    param(
        [System.IO.FileInfo]$Subtitle,
        [System.IO.FileInfo]$Video,
        [string]$LanguageTag,
        [bool]$ProcessLanguageFolders,
        [string]$FolderLanguageTag
    )

    $newName = if ($LanguageTag) {
        '{0}.{1}{2}' -f $Video.BaseName, $LanguageTag, $Subtitle.Extension
    }
    else {
        '{0}{1}' -f $Video.BaseName, $Subtitle.Extension
    }

    $targetDirectory = if ($ProcessLanguageFolders -and $FolderLanguageTag) {
        $Video.DirectoryName
    }
    else {
        $Subtitle.DirectoryName
    }

    return [PSCustomObject]@{
        NewName         = $newName
        TargetDirectory = $targetDirectory
        TargetPath      = (Join-Path -Path $targetDirectory -ChildPath $newName)
    }
}

function Write-ResultSummary {
    <#
        输出本次处理的结果汇总，便于快速查看成功/跳过/失败数量。
    #>
    param(
        [object[]]$Results
    )

    if (-not $Results -or $Results.Count -eq 0) {
        Write-Host '未产生任何处理结果。' -ForegroundColor Yellow
        return
    }

    Write-Host ''
    Write-Host '处理结果汇总：' -ForegroundColor Cyan
    foreach ($group in ($Results | Group-Object Status | Sort-Object Name)) {
        Write-Host ('  {0,-14}: {1}' -f $group.Name, $group.Count)
    }
}

#region 主流程
if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
    throw "目录不存在：$Directory"
}

$resolvedDirectory = (Resolve-Path -LiteralPath $Directory).Path
$searchPath = Join-Path -Path $resolvedDirectory -ChildPath '*'

$videoPatterns = New-SearchPatterns -Query $VideoQuery -Extensions $script:VideoExtensions
$subtitlePatterns = New-SearchPatterns -Query $SubtitleQuery -Extensions $script:SubtitleExtensions
$requestedLanguageTag = if ([string]::IsNullOrWhiteSpace($LanguageHint)) {
    $null
}
else {
    Get-LanguageTag -Name $LanguageHint -Rules $script:LanguageRules
}

# 检测是否存在 `sc / tc / jp` 等语言子文件夹；如果存在，运行中提示用户是否自动处理。
$languageFolders = Get-LanguageFolders -RootPath $resolvedDirectory -Rules $script:LanguageRules
$processLanguageFolders = $Recurse.IsPresent
if (-not $Recurse.IsPresent -and $languageFolders.Count -gt 0) {
    $folderSummary = ($languageFolders | ForEach-Object { "$($_.Name) [$($_.Language)]" }) -join ', '
    try {
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
            (New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', '自动搜索这些子文件夹，并将字幕移动到视频旁边。'),
            (New-Object System.Management.Automation.Host.ChoiceDescription '&No', '仅处理当前目录中的字幕文件。')
        )
        $selection = $Host.UI.PromptForChoice(
            '检测到语言子文件夹',
            "发现可能的字幕语言子目录：$folderSummary`n是否自动处理这些子目录中的字幕？",
            $choices,
            1
        )

        if ($selection -eq 0) {
            $processLanguageFolders = $true
        }
    }
    catch {
        Write-Warning "检测到了语言子文件夹（$folderSummary），但当前环境无法弹出交互提示。若要自动处理子目录，请改用 -Recurse。"
    }
}

$videoFiles = Get-FilesSafely -SearchPath $searchPath -Patterns $videoPatterns -Recursive $Recurse.IsPresent -Label '视频文件'
$subtitleFiles = Get-FilesSafely -SearchPath $searchPath -Patterns $subtitlePatterns -Recursive ($Recurse.IsPresent -or $processLanguageFolders) -Label '字幕文件'

if (-not $videoFiles) {
    Write-Warning "目录 '$resolvedDirectory' 中未找到视频文件。"
    return
}

if (-not $subtitleFiles) {
    Write-Warning "目录 '$resolvedDirectory' 中未找到字幕文件。"
    return
}

# 先将视频按集数分组，供后续字幕逐个匹配。
$videosByEpisode = @{}
foreach ($video in $videoFiles | Sort-Object Name) {
    $episodeKey = Get-EpisodeKey -Name $video.BaseName -Patterns $VideoEpisodePatterns
    if (-not $episodeKey) {
        Write-Verbose "跳过无法识别集数的视频：$($video.Name)"
        continue
    }

    if (-not $videosByEpisode.ContainsKey($episodeKey)) {
        $videosByEpisode[$episodeKey] = New-Object System.Collections.Generic.List[object]
    }

    $videosByEpisode[$episodeKey].Add($video)
}

if ($videosByEpisode.Count -eq 0) {
    Write-Warning "已找到视频文件，但没有任何文件能识别出有效集数。"
    return
}

# 结果表：用于回显每个字幕的处理状态
$results = New-Object System.Collections.Generic.List[object]
foreach ($subtitle in $subtitleFiles | Sort-Object Name) {
    $episodeKey = Get-EpisodeKey -Name $subtitle.BaseName -Patterns $SubtitleEpisodePatterns
    if (-not $episodeKey) {
        Write-Warning "跳过无法识别集数的字幕：$($subtitle.Name)"
        continue
    }

    if (-not $videosByEpisode.ContainsKey($episodeKey)) {
        # 如果某个视频没有字幕对应，它不会被改动；这里仅对“找不到视频的字幕”给出提示。
        Write-Warning "未找到与字幕匹配的视频：$($subtitle.Name) [$episodeKey]"
        continue
    }

    [object[]]$candidateVideos = $videosByEpisode[$episodeKey].ToArray()
    $video = Select-BestVideoMatch -Videos $candidateVideos -SubtitleName $subtitle.BaseName -Rules $script:LanguageRules
    if (-not $video) {
        $candidateNames = ($candidateVideos | Select-Object -ExpandProperty Name) -join ', '
        Write-Warning "该字幕对应多个候选视频，无法自动判断：$($subtitle.Name) [$episodeKey]。候选：$candidateNames"
        $results.Add((New-ResultObject -Status 'Ambiguous' -EpisodeKey $episodeKey -Language $null -OldName $subtitle.Name -NewName $null)) | Out-Null
        continue
    }

    # 语种优先级：文件名 > 所在语言目录 > 用户输入的 SubtitleQuery 回退提示
    $folderLanguageTag = Get-FolderLanguageTag -Path $subtitle.DirectoryName -LanguageFolders $languageFolders
    $languageTag = Get-LanguageTag -Name $subtitle.BaseName -Rules $script:LanguageRules
    if (-not $languageTag) {
        $languageTag = $folderLanguageTag
    }
    if (-not $languageTag) {
        $languageTag = $requestedLanguageTag
    }

    $targetInfo = Resolve-SubtitleTarget -Subtitle $subtitle -Video $video -LanguageTag $languageTag -ProcessLanguageFolders $processLanguageFolders -FolderLanguageTag $folderLanguageTag
    $newName = $targetInfo.NewName
    $targetDirectory = $targetInfo.TargetDirectory
    $targetPath = $targetInfo.TargetPath

    if ($subtitle.FullName -eq $targetPath) {
        $results.Add((New-ResultObject -Status 'Unchanged' -EpisodeKey $episodeKey -Language $languageTag -OldName $subtitle.Name -NewName $newName)) | Out-Null
        continue
    }

    if (Test-Path -LiteralPath $targetPath) {
        Write-Warning "目标文件已存在，已跳过：$newName"
        $results.Add((New-ResultObject -Status 'SkippedExists' -EpisodeKey $episodeKey -Language $languageTag -OldName $subtitle.Name -NewName $newName)) | Out-Null
        continue
    }

    try {
        if ($PSCmdlet.ShouldProcess($subtitle.FullName, "Move/Rename to $targetPath")) {
            # 使用 Move-Item 是因为它既能改名，也能在需要时把字幕从语言子目录移动到视频旁边。
            Move-Item -LiteralPath $subtitle.FullName -Destination $targetPath
            $status = if ($targetDirectory -ne $subtitle.DirectoryName) { 'MovedRenamed' } else { 'Renamed' }
        }
        else {
            $status = 'WhatIf'
        }
    }
    catch {
        Write-Warning "处理字幕失败：$($subtitle.Name) -> $newName，原因：$($_.Exception.Message)"
        $results.Add((New-ResultObject -Status 'Error' -EpisodeKey $episodeKey -Language $languageTag -OldName $subtitle.Name -NewName $newName)) | Out-Null
        continue
    }

    $results.Add((New-ResultObject -Status $status -EpisodeKey $episodeKey -Language $languageTag -OldName $subtitle.Name -NewName $newName)) | Out-Null
}

$results
Write-ResultSummary -Results $results
#endregion
