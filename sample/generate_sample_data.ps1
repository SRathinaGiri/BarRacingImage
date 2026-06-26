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
<ellipse cx="44" cy="28" rx="14" ry="24" fill="$Color"/>
<ellipse cx="84" cy="28" rx="14" ry="24" fill="$Color"/>
<circle cx="64" cy="70" r="42" fill="$Color"/>
<ellipse cx="54" cy="62" rx="7" ry="10" fill="#fff"/>
<ellipse cx="74" cy="62" rx="7" ry="10" fill="#fff"/>
"@
        }
        "square" {
            @"
<rect x="20" y="20" width="88" height="88" rx="14" fill="$Color"/>
<circle cx="48" cy="55" r="8" fill="#fff"/>
<circle cx="80" cy="55" r="8" fill="#fff"/>
<rect x="42" y="80" width="44" height="8" rx="4" fill="$Accent"/>
"@
        }
        "duo" {
            @"
<circle cx="48" cy="66" r="34" fill="$Color"/>
<circle cx="82" cy="66" r="34" fill="$Accent"/>
<circle cx="44" cy="58" r="6" fill="#fff"/>
<circle cx="86" cy="58" r="6" fill="#fff"/>
"@
        }
        "bolt" {
            @"
<circle cx="64" cy="64" r="44" fill="$Color"/>
<polygon points="66,22 42,70 62,70 52,106 88,56 68,56" fill="$Accent"/>
"@
        }
        default {
            @"
<circle cx="64" cy="64" r="46" fill="$Color"/>
<circle cx="64" cy="64" r="28" fill="$Accent" opacity="0.9"/>
"@
        }
    }

    @"
<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128" role="img" aria-label="$safeName">
  <rect width="128" height="128" rx="18" fill="#f8f8f8"/>
  $shapeMarkup
  <text x="64" y="76" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="24" font-weight="700" fill="#111">$safeInitials</text>
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
