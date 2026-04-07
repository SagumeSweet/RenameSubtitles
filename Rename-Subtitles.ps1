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
    [string[]]$VideoEpisodePatterns = @(),
    [string[]]$SubtitleEpisodePatterns = @(),
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

# 语种别名表：参考 ISO 639 常用语言码 + 动漫/影视字幕社区常见写法统一管理。
$script:LanguageRules = [ordered]@{
    jp = @(
        'jp', 'jpn', 'ja', 'jap', 'japanese', 'nihongo',
        '日本語', '日本语', '日語', '日语', '日文', '日配', '日字'
    )
    en = @(
        'en', 'eng', 'english',
        '英文', '英语', '英語', '英字'
    )
    sc = @(
        'sc', 'chs', 'zhs', 'zh-hans', 'zh_cn', 'zh-cn', 'cn', 'gb', 'gb2312',
        '简中', '簡中', '简体', '简体中文', '简体字幕', '简体中字'
    )
    tc = @(
        'tc', 'cht', 'zht', 'zh-hant', 'zh_tw', 'zh-tw', 'big5',
        '繁中', '繁体', '繁體', '繁体中文', '繁體中文', '繁体字幕', '繁體字幕', '繁体中字', '繁體中字'
    )
}

# 常见“发布/压制噪声”规则：参考 P2P/Scene/Servarr/TRaSH 等命名习惯统一维护。
$script:ReleaseNoiseRules = [ordered]@{
    Resolution = @(
        '\b(?:4320|2160|1440|1080|720|576|540|480)p\b',
        '\b(?:8k|4k|uhd)\b'
    )
    Source = @(
        '\b(?:web(?:[-. ]?(?:dl|rip|cap))?|bluray|blu[-. ]?ray|b[dr]rip|brrip|bdmv|bdremux|remux|hdrip|dvdrip|dvd(?:5|9)?|hdtv|tvrip)\b'
    )
    VideoCodec = @(
        '\b(?:x26[45]|xvid|divx|h\.?26[45]|hevc|avc|av1|vp9|vc-?1)\b'
    )
    BitDepth = @(
        '\b(?:hi10p|10bit|8bit)\b'
    )
    AudioCodec = @(
        '\b(?:aac(?:2\.0|5\.1)?|flac|alac|truehd|atmos|dts(?:-?hd(?:ma|hra)?|-?es|x)?|eac3|ac3|ddp(?:\+)?|dd\+|opus|mp3|pcm|lpcm)\b'
    )
    DynamicRange = @(
        '\b(?:hdr10\+?|hdr|dolby[-. ]?vision|dovi|dv|sdr)\b'
    )
    StreamingService = @(
        '\b(?:amzn|atvp|dsnp|nf|hmax|u-?next|viu|tver|fod|abema|b[- ]?global|bilibili|bahamut|baha)\b'
    )
    ReleaseMeta = @(
        '\b(?:proper|repack\d*|rerip|uncut|complete|batch|合集|内封|內封|内嵌|內嵌|外挂|外掛|dual[-. ]?audio|multi(?:[-. ]?sub(?:s)?)?|softsub(?:s)?|hardsub(?:s)?|fansub(?:s)?|official|retail|cc|sdh|dub(?:bed)?|subbed|raw(?:s)?|v[2-4])\b'
    )
}

# 默认集数识别正则：当外部传入 $null 或空数组时，自动回退到这一组。
[string[]]$script:DefaultEpisodePatterns = @(
    '(?i)\bS(?<season>\d{1,2})[ ._-]*E(?<episode>\d{1,3})\b',
    '(?i)\bE(?<episode>\d{1,3})\b',
    '(?i)\[(?<episode>\d{1,3})\]',
    '(?i)第\s*(?<episode>\d{1,3})\s*[集话話]',
    '(?i)(?:^|[\s._\-\[\(])(?<episode>\d{1,2})(?:$|[\s._\-\]\)])'
)

# 统一控制语种标签输出顺序，并缓存语种别名元数据/噪声正则，避免每次识别都重复构造。
[string[]]$script:LanguagePriority = @('jp', 'en', 'sc', 'tc')
$script:LanguageMetadataCache = $null
$script:ReleaseNoisePatternCache = $null
#endregion

function Write-Log {
    <#
        统一控制台日志输出格式，便于观察当前处理进度。
    #>
    param(
        [ValidateSet('Info', 'Step', 'Success', 'Warn')]
        [string]$Level = 'Info',
        [string]$Message
    )

    switch ($Level) {
        'Info'    { Write-Host "[info] $Message" -ForegroundColor Cyan }
        'Step'    { Write-Host "[step] $Message" -ForegroundColor DarkCyan }
        'Success' { Write-Host "[success] $Message" -ForegroundColor Green }
        'Warn'    { Write-Host "[warn] $Message" -ForegroundColor Yellow }
    }
}

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

function Resolve-PatternList {
    <#
        兼容外部传入 `$null`、空数组、空字符串的情况；
        一旦无有效正则，则自动回退到默认集数正则。
    #>
    param(
        [string[]]$Patterns,
        [string[]]$Fallback
    )

    [string[]]$resolvedPatterns = @($Patterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (@($resolvedPatterns).Count -eq 0) {
        return @($Fallback)
    }

    return $resolvedPatterns
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

function Get-ReleaseNoisePattern {
    <#
        将按类别维护的噪声规则拼成一个统一正则。
        后续如果要新增 WEB 源、音频编码、平台名，只需改顶部常量即可。
    #>
    param(
        [System.Collections.IDictionary]$Rules
    )

    if ($script:ReleaseNoisePatternCache) {
        return $script:ReleaseNoisePatternCache
    }

    [string[]]$patternParts = @(
        $Rules.GetEnumerator() | ForEach-Object {
            @($_.Value) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
    )

    if (@($patternParts).Count -eq 0) {
        return $null
    }

    $script:ReleaseNoisePatternCache = '(?i)(?:' + (($patternParts | ForEach-Object { "(?:$_)" }) -join '|') + ')'
    return $script:ReleaseNoisePatternCache
}

function Remove-ReleaseNoise {
    <#
        去掉常见发布信息噪声，减少把 1080p / x265 / 10bit 等误识别成集数的概率。
    #>
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }

    $noisePattern = Get-ReleaseNoisePattern -Rules $script:ReleaseNoiseRules
    if ([string]::IsNullOrWhiteSpace($noisePattern)) {
        return $Text
    }

    return [regex]::Replace($Text, $noisePattern, ' ')
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

        if (@($captureValues).Count -ge 1) {
            return ('E{0:D2}' -f [int]$captureValues[0])
        }
    }

    return $null
}

function Get-LanguageMetadata {
    <#
        将语种规则预处理为可复用的别名元数据，并做脚本级缓存。
        这样可避免每次识别语言时都重新遍历、排序整套别名表。
    #>
    param(
        [System.Collections.IDictionary]$Rules
    )

    if ($script:LanguageMetadataCache) {
        return $script:LanguageMetadataCache
    }

    $aliasEntries = foreach ($entry in $Rules.GetEnumerator()) {
        $canonical = ([string]$entry.Key).Trim().ToLowerInvariant()
        foreach ($alias in @($canonical) + @($entry.Value)) {
            $normalizedAlias = ([string]$alias).Trim().ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($normalizedAlias)) {
                [PSCustomObject]@{
                    Canonical = $canonical
                    Alias     = $normalizedAlias
                    Length    = $normalizedAlias.Length
                }
            }
        }
    }

    $script:LanguageMetadataCache = [PSCustomObject]@{
        AliasEntries = @($aliasEntries | Sort-Object Alias -Unique | Sort-Object Length -Descending)
        AliasNames   = @($aliasEntries | Select-Object -ExpandProperty Alias -Unique)
    }

    return $script:LanguageMetadataCache
}

function Convert-ToCanonicalLanguageTag {
    <#
        将已识别的语种标签标准化为统一格式。
        规则：统一大写、无分隔符、中文标签放最后，例如 JPSC / JPTC。
    #>
    param(
        [string[]]$Tags
    )

    [string[]]$normalizedTags = @(
        $Tags |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim().ToLowerInvariant() } |
            Select-Object -Unique
    )

    if (@($normalizedTags).Count -eq 0) {
        return $null
    }

    [string[]]$orderedTags = @()
    $orderedTags += @($normalizedTags | Where-Object { $_ -notin @('sc', 'tc') } | Sort-Object {
        $index = $script:LanguagePriority.IndexOf($_)
        if ($index -ge 0) { $index } else { 999 }
    })
    $orderedTags += @($normalizedTags | Where-Object { $_ -in @('sc', 'tc') } | Sort-Object {
        $index = $script:LanguagePriority.IndexOf($_)
        if ($index -ge 0) { $index } else { 999 }
    })

    return ((@($orderedTags | Select-Object -Unique) -join '').ToUpperInvariant())
}

function Get-LanguageTagCore {
    <#
        统一的语种识别内核：
        - 普通模式：适合文件名识别，允许在噪声词中提取语种别名。
        - 严格模式：要求整个 token 都能被语种别名完整解释，适合判断目录名是否就是语言目录。
    #>
    param(
        [string]$Name,
        [System.Collections.IDictionary]$Rules,
        [switch]$RequireCompleteMatch
    )

    $normalizedName = ([string]$Name).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($normalizedName)) {
        return $null
    }

    $metadata = Get-LanguageMetadata -Rules $Rules
    $detectedTags = New-Object System.Collections.Generic.List[string]

    if (-not $RequireCompleteMatch) {
        foreach ($entry in $metadata.AliasEntries) {
            $escapedAlias = [regex]::Escape($entry.Alias)
            $boundaryPattern = "(?i)(?:^|[\.\-_\s\[\]\(\)\{\}])$escapedAlias(?:$|[\.\-_\s\[\]\(\)\{\}])"
            if ($normalizedName -match $boundaryPattern -and -not $detectedTags.Contains($entry.Canonical)) {
                $detectedTags.Add($entry.Canonical) | Out-Null
            }
        }
    }

    $tokens = @([regex]::Split($normalizedName, '[\.\-_\s\[\]\(\)\{\}]+') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (@($tokens).Count -eq 0) {
        return $null
    }

    foreach ($token in $tokens) {
        $rest = $token
        $tokenMatches = New-Object System.Collections.Generic.List[string]

        while ($rest.Length -gt 0) {
            $matchedAlias = $metadata.AliasEntries | Where-Object { $rest.StartsWith($_.Alias) } | Select-Object -First 1
            if (-not $matchedAlias) {
                if ($RequireCompleteMatch) {
                    return $null
                }

                $tokenMatches.Clear()
                break
            }

            if (-not $tokenMatches.Contains($matchedAlias.Canonical)) {
                $tokenMatches.Add($matchedAlias.Canonical) | Out-Null
            }
            $rest = $rest.Substring($matchedAlias.Alias.Length)
        }

        if ($rest.Length -eq 0) {
            foreach ($tag in $tokenMatches) {
                if (-not $detectedTags.Contains($tag)) {
                    $detectedTags.Add($tag) | Out-Null
                }
            }
        }
    }

    return Convert-ToCanonicalLanguageTag -Tags $detectedTags
}

function Get-LanguageTag {
    <#
        从文件名中识别语种标签。
        支持分隔形式（.jp / -tc / _chs）以及紧凑组合（SCJP / BIG5JP）。
    #>
    param(
        [string]$Name,
        [System.Collections.IDictionary]$Rules
    )

    return Get-LanguageTagCore -Name $Name -Rules $Rules
}

function Get-ExactLanguageTag {
    <#
        用于“语言子目录自动发现”的严格匹配函数。
        只有当整个目录名都能被语种别名完整解释时，才认定它是语言目录。
    #>
    param(
        [string]$Name,
        [System.Collections.IDictionary]$Rules
    )

    return Get-LanguageTagCore -Name $Name -Rules $Rules -RequireCompleteMatch
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

    $languageMetadata = Get-LanguageMetadata -Rules $Rules
    [string[]]$languageAliases = @($languageMetadata.AliasNames)

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
    if (@($videoArray).Count -eq 0) {
        return $null
    }

    if (@($videoArray).Count -eq 1) {
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

    if (@($orderedCandidates).Count -eq 1) {
        return $orderedCandidates[0].Video
    }

    if (@($orderedCandidates).Count -ge 2 -and $orderedCandidates[0].Score -gt $orderedCandidates[1].Score) {
        return $orderedCandidates[0].Video
    }

    return $null
}

function Convert-ToLanguageFolderInfo {
    <#
        将语言目录候选项统一转换为标准结构，避免不同对象形态导致 `.Name/.Path/.Language`
        属性访问失败。
    #>
    param(
        [object]$InputObject,
        [System.Collections.IDictionary]$Rules
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $name = $null
    $path = $null
    $language = $null

    if ($InputObject -is [string]) {
        $name = $InputObject.Trim()
    }
    else {
        $properties = $InputObject.PSObject.Properties

        if ($properties.Match('Name').Count -gt 0 -and $null -ne $InputObject.Name) {
            $name = [string]$InputObject.Name
        }
        elseif ($properties.Match('BaseName').Count -gt 0 -and $null -ne $InputObject.BaseName) {
            $name = [string]$InputObject.BaseName
        }

        if ($properties.Match('Path').Count -gt 0 -and $null -ne $InputObject.Path) {
            $path = [string]$InputObject.Path
        }
        elseif ($properties.Match('FullName').Count -gt 0 -and $null -ne $InputObject.FullName) {
            $path = [string]$InputObject.FullName
        }

        if ($properties.Match('Language').Count -gt 0 -and $null -ne $InputObject.Language) {
            $language = [string]$InputObject.Language
        }

        if ([string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($path)) {
            $name = Split-Path -Leaf $path
        }
    }

    if (-not $language -and -not [string]::IsNullOrWhiteSpace($name)) {
        $language = Get-ExactLanguageTag -Name $name -Rules $Rules
    }

    if ([string]::IsNullOrWhiteSpace($language)) {
        return $null
    }

    return [PSCustomObject]@{
        Name     = $name
        Path     = $path
        Language = $language.ToUpperInvariant()
    }
}

function Get-LanguageFolderSummary {
    <#
        将语言目录列表格式化为稳定的摘要文本，避免直接访问不存在的属性。
    #>
    param(
        [object[]]$LanguageFolders,
        [System.Collections.IDictionary]$Rules
    )

    [string[]]$items = @(
        $LanguageFolders | ForEach-Object {
            $folderInfo = Convert-ToLanguageFolderInfo -InputObject $_ -Rules $Rules
            if ($null -ne $folderInfo) {
                if ([string]::IsNullOrWhiteSpace($folderInfo.Name)) {
                    "[$($folderInfo.Language)]"
                }
                else {
                    "$($folderInfo.Name) [$($folderInfo.Language)]"
                }
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    return ($items -join ', ')
}

function Get-FolderLanguageTag {
    <#
        如果字幕位于语言子文件夹中（如 sc / tc / jp），则从路径继承语言标签。
    #>
    param(
        [string]$Path,
        [object[]]$LanguageFolders
    )

    [object[]]$normalizedFolders = @(
        $LanguageFolders |
            ForEach-Object { Convert-ToLanguageFolderInfo -InputObject $_ -Rules $script:LanguageRules } |
            Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace($_.Path) } |
            Sort-Object { $_.Path.Length } -Descending
    )

    foreach ($folder in $normalizedFolders) {
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
        foreach ($directory in @(Get-ChildItem -LiteralPath $RootPath -Directory -ErrorAction SilentlyContinue)) {
            $folderInfo = Convert-ToLanguageFolderInfo -InputObject $directory -Rules $Rules
            if ($folderInfo) {
                $folderInfo
            }
        }
    )
}

function Test-AnyPatternMatch {
    <#
        判断文件名是否命中任意一个 PowerShell 通配符模式。
    #>
    param(
        [string]$Name,
        [string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if ($Name -like $pattern) {
            return $true
        }
    }

    return $false
}

function Get-FilesSafely {
    <#
        统一封装文件枚举，出错时抛出更明确的错误信息。
        使用 -LiteralPath，避免目录名中包含 [] 等字符时被当成通配符解释。
    #>
    param(
        [string]$SearchPath,
        [string[]]$Patterns,
        [bool]$Recursive,
        [string]$Label
    )

    try {
        $items = if ($Recursive) {
            Get-ChildItem -LiteralPath $SearchPath -File -Recurse -ErrorAction Stop
        }
        else {
            Get-ChildItem -LiteralPath $SearchPath -File -ErrorAction Stop
        }

        return @($items | Where-Object { Test-AnyPatternMatch -Name $_.Name -Patterns $Patterns })
    }
    catch {
        throw "枚举${Label}失败：$($_.Exception.Message)"
    }
}

function Confirm-LanguageFolderProcessing {
    <#
        将“是否自动处理语言子目录”的交互逻辑独立出来，减少主流程分支复杂度。
        这里会先对目录对象做标准化，避免远程环境中出现对象属性不一致的问题。
    #>
    param(
        [object[]]$LanguageFolders,
        [bool]$DefaultValue,
        [bool]$IsRecursive
    )

    [object[]]$normalizedFolders = @(
        $LanguageFolders |
            ForEach-Object { Convert-ToLanguageFolderInfo -InputObject $_ -Rules $script:LanguageRules } |
            Where-Object { $_ }
    )

    if ($IsRecursive -or @($normalizedFolders).Count -eq 0) {
        return $DefaultValue
    }

    $folderSummary = Get-LanguageFolderSummary -LanguageFolders $normalizedFolders -Rules $script:LanguageRules
    if ([string]::IsNullOrWhiteSpace($folderSummary)) {
        return $DefaultValue
    }

    Write-Log -Level Info -Message "检测到可能的语言子目录：$folderSummary"

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
            Write-Log -Level Success -Message '已启用语言子目录自动处理。'
            return $true
        }

        Write-Log -Level Info -Message '将仅处理当前目录中的字幕文件。'
        return $false
    }
    catch {
        Write-Log -Level Warn -Message "检测到了语言子文件夹（$folderSummary），但当前环境无法弹出交互提示。若要自动处理子目录，请改用 -Recurse。"
        return $DefaultValue
    }
}

function Get-VideoEpisodeIndex {
    <#
        将视频预先建立为“集数键 -> 候选视频列表”的索引，便于后续字幕匹配。
    #>
    param(
        [System.IO.FileInfo[]]$Videos,
        [string[]]$EpisodePatterns
    )

    $videosByEpisode = @{}
    foreach ($video in @($Videos | Sort-Object Name)) {
        $episodeKey = Get-EpisodeKey -Name $video.BaseName -Patterns $EpisodePatterns
        if (-not $episodeKey) {
            Write-Verbose "跳过无法识别集数的视频：$($video.Name)"
            continue
        }

        if (-not $videosByEpisode.ContainsKey($episodeKey)) {
            $videosByEpisode[$episodeKey] = New-Object System.Collections.Generic.List[object]
        }

        $videosByEpisode[$episodeKey].Add($video)
    }

    return $videosByEpisode
}

function Resolve-SubtitleLanguage {
    <#
        统一语种决策优先级：文件名 > 所在语言目录 > 用户传入的 LanguageHint。
    #>
    param(
        [System.IO.FileInfo]$Subtitle,
        [object[]]$LanguageFolders,
        [string]$RequestedLanguageTag,
        [System.Collections.IDictionary]$Rules
    )

    $folderLanguageTag = Get-FolderLanguageTag -Path $Subtitle.DirectoryName -LanguageFolders $LanguageFolders
    $fileLanguageTag = Get-LanguageTag -Name $Subtitle.BaseName -Rules $Rules

    $resolvedLanguageTag = if ($fileLanguageTag) {
        $fileLanguageTag
    }
    elseif ($folderLanguageTag) {
        $folderLanguageTag
    }
    else {
        $RequestedLanguageTag
    }

    return [PSCustomObject]@{
        FileLanguageTag   = $fileLanguageTag
        FolderLanguageTag = $folderLanguageTag
        LanguageTag       = $resolvedLanguageTag
    }
}

function Invoke-SubtitleRenameAction {
    <#
        处理单个字幕文件：识别集数、匹配视频、解析语种、执行改名/移动。
        将单文件逻辑收敛到一个函数中，主流程只负责“扫描与调度”。
    #>
    param(
        [System.Management.Automation.PSCmdlet]$Cmdlet,
        [System.IO.FileInfo]$Subtitle,
        [hashtable]$VideosByEpisode,
        [string[]]$EpisodePatterns,
        [object[]]$LanguageFolders,
        [string]$RequestedLanguageTag,
        [bool]$ProcessLanguageFolders,
        [System.Collections.IDictionary]$Rules
    )

    Write-Log -Level Step -Message "处理字幕：$($Subtitle.Name)"

    $episodeKey = Get-EpisodeKey -Name $Subtitle.BaseName -Patterns $EpisodePatterns
    if (-not $episodeKey) {
        Write-Log -Level Warn -Message "无法识别集数，已跳过：$($Subtitle.Name)"
        return $null
    }

    Write-Log -Level Info -Message "识别到集数键：$episodeKey"

    if (-not $VideosByEpisode.ContainsKey($episodeKey)) {
        Write-Log -Level Warn -Message "未找到对应视频：$($Subtitle.Name) [$episodeKey]"
        return $null
    }

    [object[]]$candidateVideos = $VideosByEpisode[$episodeKey].ToArray()
    $video = Select-BestVideoMatch -Videos $candidateVideos -SubtitleName $Subtitle.BaseName -Rules $Rules
    if (-not $video) {
        $candidateNames = ($candidateVideos | Select-Object -ExpandProperty Name) -join ', '
        Write-Log -Level Warn -Message "存在多个候选视频，无法自动判断：$($Subtitle.Name) [$episodeKey]。候选：$candidateNames"
        return (New-ResultObject -Status 'Ambiguous' -EpisodeKey $episodeKey -Language $null -OldName $Subtitle.Name -NewName $null)
    }

    Write-Log -Level Info -Message "匹配视频：$($video.Name)"

    $languageDecision = Resolve-SubtitleLanguage -Subtitle $Subtitle -LanguageFolders $LanguageFolders -RequestedLanguageTag $RequestedLanguageTag -Rules $Rules
    $languageTag = $languageDecision.LanguageTag
    $targetInfo = Resolve-SubtitleTarget -Subtitle $Subtitle -Video $video -LanguageTag $languageTag -ProcessLanguageFolders $ProcessLanguageFolders -FolderLanguageTag $languageDecision.FolderLanguageTag

    if ($languageTag) {
        Write-Log -Level Info -Message "识别语种：$languageTag"
    }
    else {
        Write-Log -Level Info -Message '未识别到语种标签，将直接使用视频主体名。'
    }

    Write-Log -Level Info -Message "目标路径：$($targetInfo.TargetPath)"

    if ($Subtitle.FullName -eq $targetInfo.TargetPath) {
        Write-Log -Level Info -Message '文件名已符合目标格式，无需修改。'
        return (New-ResultObject -Status 'Unchanged' -EpisodeKey $episodeKey -Language $languageTag -OldName $Subtitle.Name -NewName $targetInfo.NewName)
    }

    if (Test-Path -LiteralPath $targetInfo.TargetPath) {
        Write-Log -Level Warn -Message "目标文件已存在，已跳过：$($targetInfo.NewName)"
        return (New-ResultObject -Status 'SkippedExists' -EpisodeKey $episodeKey -Language $languageTag -OldName $Subtitle.Name -NewName $targetInfo.NewName)
    }

    try {
        if ($Cmdlet.ShouldProcess($Subtitle.FullName, "Move/Rename to $($targetInfo.TargetPath)")) {
            Move-Item -LiteralPath $Subtitle.FullName -Destination $targetInfo.TargetPath -ErrorAction Stop
            $status = if ($targetInfo.TargetDirectory -ne $Subtitle.DirectoryName) { 'MovedRenamed' } else { 'Renamed' }
            Write-Log -Level Success -Message "已处理：$($Subtitle.Name) -> $($targetInfo.NewName)"
        }
        else {
            $status = 'WhatIf'
            Write-Log -Level Info -Message "预览模式：将执行 $($Subtitle.Name) -> $($targetInfo.NewName)"
        }
    }
    catch {
        Write-Log -Level Warn -Message "处理字幕失败：$($Subtitle.Name) -> $($targetInfo.NewName)，原因：$($_.Exception.Message)"
        return (New-ResultObject -Status 'Error' -EpisodeKey $episodeKey -Language $languageTag -OldName $Subtitle.Name -NewName $targetInfo.NewName)
    }

    return (New-ResultObject -Status $status -EpisodeKey $episodeKey -Language $languageTag -OldName $Subtitle.Name -NewName $targetInfo.NewName)
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

    if (-not $Results -or @($Results).Count -eq 0) {
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
$searchPath = $resolvedDirectory

Write-Log -Level Step -Message "开始处理目录：$resolvedDirectory"
Write-Log -Level Info -Message "视频筛选关键字：$([string]::IsNullOrWhiteSpace($VideoQuery) ? '<空>' : $VideoQuery)"
Write-Log -Level Info -Message "字幕筛选关键字：$([string]::IsNullOrWhiteSpace($SubtitleQuery) ? '<空>' : $SubtitleQuery)"
Write-Log -Level Info -Message "是否递归扫描：$($Recurse.IsPresent)"

$videoPatterns = New-SearchPatterns -Query $VideoQuery -Extensions $script:VideoExtensions
$subtitlePatterns = New-SearchPatterns -Query $SubtitleQuery -Extensions $script:SubtitleExtensions
$VideoEpisodePatterns = Resolve-PatternList -Patterns $VideoEpisodePatterns -Fallback $script:DefaultEpisodePatterns
$SubtitleEpisodePatterns = Resolve-PatternList -Patterns $SubtitleEpisodePatterns -Fallback $script:DefaultEpisodePatterns
$requestedLanguageTag = if ([string]::IsNullOrWhiteSpace($LanguageHint)) {
    $null
}
else {
    Get-LanguageTag -Name $LanguageHint -Rules $script:LanguageRules
}

if ($requestedLanguageTag) {
    Write-Log -Level Info -Message "已设置语种提示：$requestedLanguageTag"
}

# 检测是否存在 `sc / tc / jp` 等语言子文件夹；如有需要，运行时询问是否纳入处理范围。
$languageFolders = Get-LanguageFolders -RootPath $resolvedDirectory -Rules $script:LanguageRules
$processLanguageFolders = Confirm-LanguageFolderProcessing -LanguageFolders $languageFolders -DefaultValue $Recurse.IsPresent -IsRecursive $Recurse.IsPresent

$videoFiles = Get-FilesSafely -SearchPath $searchPath -Patterns $videoPatterns -Recursive $Recurse.IsPresent -Label '视频文件'
$subtitleFiles = Get-FilesSafely -SearchPath $searchPath -Patterns $subtitlePatterns -Recursive ($Recurse.IsPresent -or $processLanguageFolders) -Label '字幕文件'

if (-not $videoFiles) {
    Write-Log -Level Warn -Message "目录中未找到视频文件：$resolvedDirectory"
    return
}

if (-not $subtitleFiles) {
    Write-Log -Level Warn -Message "目录中未找到字幕文件：$resolvedDirectory"
    return
}

Write-Log -Level Info -Message "扫描完成：视频 $(@($videoFiles).Count) 个，字幕 $(@($subtitleFiles).Count) 个。"

# 先将视频按集数分组，供后续字幕逐个匹配。
$videosByEpisode = Get-VideoEpisodeIndex -Videos $videoFiles -EpisodePatterns $VideoEpisodePatterns

if (@($videosByEpisode.Keys).Count -eq 0) {
    Write-Log -Level Warn -Message '没有任何视频能识别出有效集数，无法继续匹配。'
    return
}

Write-Log -Level Success -Message "视频索引建立完成：共识别到 $(@($videosByEpisode.Keys).Count) 个集数键。"

# 结果表：用于回显每个字幕的处理状态
$results = New-Object System.Collections.Generic.List[object]
foreach ($subtitle in $subtitleFiles | Sort-Object Name) {
    $result = Invoke-SubtitleRenameAction `
        -Cmdlet $PSCmdlet `
        -Subtitle $subtitle `
        -VideosByEpisode $videosByEpisode `
        -EpisodePatterns $SubtitleEpisodePatterns `
        -LanguageFolders $languageFolders `
        -RequestedLanguageTag $requestedLanguageTag `
        -ProcessLanguageFolders $processLanguageFolders `
        -Rules $script:LanguageRules

    if ($result) {
        $results.Add($result) | Out-Null
    }
}

$results
Write-ResultSummary -Results $results
#endregion
