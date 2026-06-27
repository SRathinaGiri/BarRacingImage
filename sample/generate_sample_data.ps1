param(
    [string]$OutputPath = ".\sample\BarRacingImageSampleData.csv",
    [string]$ImageFolder = ".\sample\images",
    [int]$StartYear = 2024,
    [int]$EndYear = 2024
)

$ErrorActionPreference = "Stop"

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputPath) | Out-Null
New-Item -ItemType Directory -Force -Path $ImageFolder | Out-Null

$characters = @(
    @{ Name = "Mickey Mouse"; Page = "Mickey_Mouse"; Color = "#D92332"; Accent = "#FFD23F"; Initials = "MM"; Shape = "circle" },
    @{ Name = "Bugs Bunny"; Page = "Bugs_Bunny"; Color = "#6C7A89"; Accent = "#F7C9C9"; Initials = "BB"; Shape = "ears" },
    @{ Name = "SpongeBob SquarePants"; Page = "SpongeBob_SquarePants"; Color = "#F7D843"; Accent = "#2B7BBB"; Initials = "SB"; Shape = "square" },
    @{ Name = "Doraemon"; Page = "Doraemon"; Color = "#2F9BEA"; Accent = "#FFFFFF"; Initials = "DO"; Shape = "circle" },
    @{ Name = "Tom and Jerry"; Page = "Tom_and_Jerry"; Color = "#8B5E3C"; Accent = "#F2B166"; Initials = "TJ"; Shape = "duo" },
    @{ Name = "Pikachu"; Page = "Pikachu"; Color = "#FFD83D"; Accent = "#E84545"; Initials = "PK"; Shape = "bolt" }
)

function New-IconSvg {
    param(
        [string]$Name,
        [string]$Color,
        [string]$Accent,
        [string]$Initials,
        [string]$Shape
    )

    $safeName = [System.Security.SecurityElement]::Escape($Name)
    $safeInitials = [System.Security.SecurityElement]::Escape($Initials)

    $shapeMarkup = switch ($Shape) {
        "ears" {
            @"
<ellipse cx="45" cy="29" rx="13" ry="25" fill="$Color" transform="rotate(-10 45 29)"/>
<ellipse cx="83" cy="29" rx="13" ry="25" fill="$Color" transform="rotate(10 83 29)"/>
<ellipse cx="45" cy="30" rx="6" ry="17" fill="$Accent" transform="rotate(-10 45 30)"/>
<ellipse cx="83" cy="30" rx="6" ry="17" fill="$Accent" transform="rotate(10 83 30)"/>
<circle cx="64" cy="72" r="39" fill="$Color"/>
<ellipse cx="52" cy="65" rx="8" ry="10" fill="#fff"/>
<ellipse cx="76" cy="65" rx="8" ry="10" fill="#fff"/>
<circle cx="54" cy="67" r="3" fill="#111"/>
<circle cx="74" cy="67" r="3" fill="#111"/>
<ellipse cx="64" cy="83" rx="18" ry="13" fill="#fff"/>
<rect x="58" y="82" width="5" height="12" fill="#fff" stroke="#333" stroke-width="1"/>
<rect x="65" y="82" width="5" height="12" fill="#fff" stroke="#333" stroke-width="1"/>
"@
        }
        "square" {
            @"
<rect x="22" y="16" width="84" height="96" rx="10" fill="$Color" stroke="#D7A700" stroke-width="3"/>
<circle cx="47" cy="53" r="13" fill="#fff"/>
<circle cx="81" cy="53" r="13" fill="#fff"/>
<circle cx="50" cy="54" r="5" fill="$Accent"/>
<circle cx="78" cy="54" r="5" fill="$Accent"/>
<circle cx="42" cy="78" r="3" fill="#C39C00"/>
<circle cx="88" cy="77" r="3" fill="#C39C00"/>
<rect x="46" y="86" width="36" height="8" rx="4" fill="#8B4513"/>
<rect x="51" y="93" width="8" height="10" fill="#fff"/>
<rect x="69" y="93" width="8" height="10" fill="#fff"/>
"@
        }
        "duo" {
            @"
<circle cx="46" cy="68" r="33" fill="$Color"/>
<polygon points="25,39 38,55 48,36" fill="$Color"/>
<polygon points="64,36 55,55 75,43" fill="$Color"/>
<circle cx="42" cy="61" r="8" fill="#fff"/>
<circle cx="55" cy="61" r="8" fill="#fff"/>
<circle cx="43" cy="62" r="3" fill="#111"/>
<circle cx="54" cy="62" r="3" fill="#111"/>
<ellipse cx="49" cy="78" rx="12" ry="8" fill="#EBC7A0"/>
<circle cx="88" cy="75" r="24" fill="$Accent"/>
<circle cx="80" cy="69" r="6" fill="#fff"/>
<circle cx="96" cy="69" r="6" fill="#fff"/>
<circle cx="80" cy="69" r="2" fill="#111"/>
<circle cx="96" cy="69" r="2" fill="#111"/>
<ellipse cx="88" cy="84" rx="10" ry="5" fill="#fff"/>
"@
        }
        "bolt" {
            @"
<polygon points="42,25 53,56 30,42" fill="$Color" stroke="#222" stroke-width="4"/>
<polygon points="86,25 75,56 98,42" fill="$Color" stroke="#222" stroke-width="4"/>
<circle cx="64" cy="70" r="39" fill="$Color" stroke="#222" stroke-width="3"/>
<circle cx="49" cy="66" r="7" fill="#111"/>
<circle cx="79" cy="66" r="7" fill="#111"/>
<circle cx="43" cy="82" r="9" fill="$Accent"/>
<circle cx="85" cy="82" r="9" fill="$Accent"/>
<path d="M54 90 Q64 98 74 90" fill="none" stroke="#111" stroke-width="4" stroke-linecap="round"/>
<polygon points="66,22 42,70 62,70 52,106 88,56 68,56" fill="#FFB300" opacity=".35"/>
"@
        }
        default {
            @"
<circle cx="43" cy="38" r="19" fill="#111"/>
<circle cx="85" cy="38" r="19" fill="#111"/>
<circle cx="64" cy="66" r="42" fill="$Color"/>
<ellipse cx="64" cy="72" rx="30" ry="28" fill="$Accent" opacity="0.95"/>
<circle cx="53" cy="62" r="6" fill="#111"/>
<circle cx="75" cy="62" r="6" fill="#111"/>
<ellipse cx="64" cy="78" rx="10" ry="7" fill="#111"/>
<path d="M54 90 Q64 98 74 90" fill="none" stroke="#111" stroke-width="4" stroke-linecap="round"/>
"@
        }
    }

    @"
<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128" role="img" aria-label="$safeName">
  <rect width="128" height="128" rx="18" fill="#f8f8f8"/>
  $shapeMarkup
</svg>
"@
}

function ConvertTo-DataUri {
    param([string]$Text)

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    "data:image/svg+xml;base64,$([Convert]::ToBase64String($bytes))"
}

$start = "{0}0101" -f $StartYear
$end = "{0}1231" -f $EndYear
$rows = New-Object System.Collections.Generic.List[object]

foreach ($character in $characters) {
    $svg = New-IconSvg -Name $character.Name -Color $character.Color -Accent $character.Accent -Initials $character.Initials -Shape $character.Shape
    $imageFile = Join-Path $ImageFolder (($character.Page) + ".svg")
    Set-Content -Path $imageFile -Value $svg -Encoding UTF8
    $dataUri = ConvertTo-DataUri -Text $svg

    $encodedPage = [uri]::EscapeDataString($character.Page)
    $url = "https://wikimedia.org/api/rest_v1/metrics/pageviews/per-article/en.wikipedia/all-access/user/$encodedPage/monthly/$start/$end"
    $response = Invoke-RestMethod -Uri $url -Headers @{ "User-Agent" = "BarRacingImageSample/1.0 (srg@rathinagiri.in)" }

    foreach ($item in $response.items) {
        $timestamp = [string]$item.timestamp
        $period = "{0}-{1}" -f $timestamp.Substring(0, 4), $timestamp.Substring(4, 2)
        $rows.Add([pscustomobject]@{
            Period = $period
            Category = $character.Name
            ImageURI = $dataUri
            PageViews = [int]$item.views
            Source = "Wikimedia Pageviews API"
            Article = "https://en.wikipedia.org/wiki/$($character.Page)"
        })
    }
}

$rows |
    Sort-Object Period, Category |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Created $OutputPath with $($rows.Count) rows."
Write-Host "Created SVG image assets in $ImageFolder."
