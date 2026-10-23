$ErrorActionPreference = "Stop"

$workspaceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestPath = Join-Path $workspaceRoot "manifest.json"
$manifest = Get-Content -Raw $manifestPath | ConvertFrom-Json
$packagePath = Join-Path $env:TEMP ("swe-outlook-owa-" + [Guid]::NewGuid().ToString("N") + ".zip")
$localBaseUrl = "https://localhost:3000"

function Set-LocalManifestUrls {
    param([object]$Value)

    if ($null -eq $Value) {
        return
    }

    if ($Value -is [System.Array]) {
        foreach ($item in $Value) {
            Set-LocalManifestUrls $item
        }
        return
    }

    if ($Value -isnot [pscustomobject]) {
        return
    }

    foreach ($property in $Value.PSObject.Properties) {
        if ($property.Name -in @("page", "url") -and $property.Value -is [string]) {
            $parsedUrl = $null
            if ([Uri]::TryCreate($property.Value, [UriKind]::Absolute, [ref]$parsedUrl)) {
                $path = $parsedUrl.PathAndQuery + $parsedUrl.Fragment
            }
            else {
                $path = "/" + $property.Value.TrimStart("/")
            }

            $property.Value = "$localBaseUrl$path"
        }
        else {
            Set-LocalManifestUrls $property.Value
        }
    }
}

try {
    if (-not $manifest.webApplicationInfo.id -or -not $manifest.webApplicationInfo.resource) {
        throw "The manifest must contain webApplicationInfo.id and webApplicationInfo.resource."
    }

    Set-LocalManifestUrls $manifest
    $manifest.validDomains = @("localhost")
    $manifest.webApplicationInfo.resource = "api://localhost:3000/$($manifest.webApplicationInfo.id)"
    $manifestJson = ConvertTo-Json -InputObject $manifest -Depth 100
    if ($manifestJson -match "azurestaticapps\.net") {
        throw "The local sideload package still contains an Azure Static Web Apps URL."
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::Open(
        $packagePath,
        [System.IO.Compression.ZipArchiveMode]::Create
    )
    try {
        $manifestEntry = $archive.CreateEntry(
            "manifest.json",
            [System.IO.Compression.CompressionLevel]::Optimal
        )
        $manifestWriter = [System.IO.StreamWriter]::new(
            $manifestEntry.Open(),
            [System.Text.UTF8Encoding]::new($false)
        )
        try {
            $manifestWriter.Write($manifestJson)
        }
        finally {
            $manifestWriter.Dispose()
        }

        foreach ($iconPath in @($manifest.icons.color, $manifest.icons.outline)) {
        if ([string]::IsNullOrWhiteSpace($iconPath)) {
            throw "The manifest must reference both color and outline icons."
        }

        $sourcePath = Join-Path $workspaceRoot $iconPath
        if (-not (Test-Path -LiteralPath $sourcePath)) {
            throw "Manifest icon not found: $iconPath"
        }

            $entryPath = $iconPath -replace '\\', '/'
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive,
                $sourcePath,
                $entryPath,
                [System.IO.Compression.CompressionLevel]::Optimal
            ) | Out-Null
        }
    }
    finally {
        $archive.Dispose()
    }

    Push-Location $workspaceRoot
    try {
        & npx --yes -p @microsoft/m365agentstoolkit-cli atk install --file-path $packagePath --interactive false
        if ($LASTEXITCODE -ne 0) {
            throw "Microsoft 365 app package installation failed with exit code $LASTEXITCODE."
        }
    }
    finally {
        Pop-Location
    }

    Start-Process "https://outlook.office.com/mail/"
}
finally {
    if (Test-Path -LiteralPath $packagePath) {
        Remove-Item -LiteralPath $packagePath -Force
    }
}