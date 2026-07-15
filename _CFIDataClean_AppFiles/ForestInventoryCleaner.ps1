# DADA - Database Dad
# Runs best from the 32-bit launcher so Access drivers are visible.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ([System.Management.Automation.PSTypeName]"DadaWin32.TaskbarIdentity").Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace DadaWin32
{
    public static class TaskbarIdentity
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SetCurrentProcessExplicitAppUserModelID(string appID);
    }
}
"@
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$ErrorActionPreference = "Stop"

$script:DbPath = $null
$script:ConnectionString = $null
$script:ColumnsByTable = @{}
$script:FieldManualTipsBySheet = $null
$script:AiDefaultEndpoint = "https://doi-training-foundry.services.ai.azure.com/"
$script:AiDefaultModel = "gpt-4.1-mini"
$script:AzureOpenAiApiVersion = "2024-10-21"
$script:AzureFoundryModelsApiVersion = "2024-05-01-preview"
$script:AiApiKey = ""
$script:ApprovedAiEndpointHosts = @(
    "doi-training-foundry.services.ai.azure.com",
    "doi-training-foundry.cognitiveservices.azure.com"
)
$script:AppTempFolderName = "DADA-DatabaseDad"
$script:ProjectManualPath = ""
$script:ProjectManualText = ""
$script:ProjectManualLoadedFrom = ""
$script:ProjectManualDigestByKey = @{}
$script:AiGuidanceCache = @{}
$script:AppVersion = "1.0.00"
$script:AppName = "CFI DataClean"
$script:AppUserModelId = "USDI.BIA.CFIDataClean"
$script:AppWindowBaseTitle = "$script:AppName - Build $script:AppVersion"
$script:AppIcon = $null
$script:DefaultWoodlandSpeciesCodes = ""
$script:DefaultHeightRequiredMinorPlots = ""
$script:DefaultRegenTimberSeedlingMinorPlots = ""
$script:DefaultRegenTimberSapling20MinorPlots = ""
$script:DefaultRegenTimberSapling40MinorPlots = ""
$script:DefaultRegenWoodlandSeedlingMinorPlots = ""
$script:DefaultRegenWoodlandSapling20MinorPlots = ""
$script:DefaultRegenWoodlandSapling40MinorPlots = ""
$script:GuideRunSettings = $null
$script:GuideRunWorker = $null
$script:GuideRunStartedAt = $null
$script:GuideRunCancellationRequested = $false
$script:GuideRunProgressPath = ""
$script:GuideRunCancelPath = ""
$script:GuideRunLogPath = ""
$script:LastGuideRunLoggedProgress = ""
$script:GuideRunPerformanceEntries = New-Object System.Collections.Generic.List[object]
$script:GuideRunSlowSqlEntries = New-Object System.Collections.Generic.List[object]
$script:GuideRunCurrentPerformanceStep = ""
$script:GuideRunPerformanceSummaryWritten = $false
$script:GuideRunSlowSqlThresholdSeconds = 2.0
$script:SuppressGuideRunCompletionDialogs = $false
$script:CloseAfterGuideRunCancel = $false
$script:PeriodScopeConditionCache = @{}
$script:LastAiSuggestedEditGroupsApplied = 0
$script:LastAiSuggestedEditGroupsSkipped = 0
$script:LastAiSuggestedEditRowsApplied = 0
$script:LastAiSuggestedEditRowsSkipped = 0
$script:LastAiSuggestedEditManualGroupsUsed = 0
$script:LastAiSuggestedEditManualRowsUsed = 0
$script:LastAiSuggestedEditNoManualRetryGroups = 0
$script:LastAiSuggestedEditNoManualRetryRows = 0
$script:LastAiSuggestedEditNoManualSnippetGroups = 0
$script:LastAiSuggestedEditNoManualSnippetRows = 0
$script:LastAiExportGuidanceAttemptLabel = ""
$script:AppColumnReviewMetadataCache = @{}
$script:LastReviewedCleanExportRows = 0

function Quote-Name {
    param([string]$Name)
    return "[" + $Name.Replace("]", "]]") + "]"
}

function Sql-Text {
    param([string]$Value)
    if ($null -eq $Value) { return "Null" }
    return "'" + $Value.Replace("'", "''") + "'"
}

function New-RoundedRectanglePath {
    param(
        [System.Drawing.Rectangle]$Rectangle,
        [int]$Radius
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $diameter = $Radius * 2
    $right = $Rectangle.Right - $diameter - 1
    $bottom = $Rectangle.Bottom - $diameter - 1
    $path.AddArc($Rectangle.X, $Rectangle.Y, $diameter, $diameter, 180, 90)
    $path.AddArc($right, $Rectangle.Y, $diameter, $diameter, 270, 90)
    $path.AddArc($right, $bottom, $diameter, $diameter, 0, 90)
    $path.AddArc($Rectangle.X, $bottom, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-DadaAppIcon {
    $iconPath = Join-Path $PSScriptRoot "DADA-broom.ico"
    if (Test-Path -LiteralPath $iconPath) {
        $stream = $null
        $loadedIcon = $null
        try {
            $stream = [System.IO.File]::OpenRead($iconPath)
            $loadedIcon = New-Object System.Drawing.Icon -ArgumentList $stream
            return $loadedIcon.Clone()
        }
        catch {
        }
        finally {
            if ($null -ne $loadedIcon) { $loadedIcon.Dispose() }
            if ($null -ne $stream) { $stream.Dispose() }
        }
    }

    $bitmap = New-Object System.Drawing.Bitmap -ArgumentList 64, 64
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $rect = New-Object System.Drawing.Rectangle -ArgumentList 2, 2, 60, 60
        $path = New-RoundedRectanglePath -Rectangle $rect -Radius 14
        $startColor = [System.Drawing.Color]::FromArgb(22, 92, 104)
        $endColor = [System.Drawing.Color]::FromArgb(52, 135, 83)
        $background = New-Object System.Drawing.Drawing2D.LinearGradientBrush -ArgumentList $rect, $startColor, $endColor, 45.0
        try {
            $graphics.FillPath($background, $path)
        }
        finally {
            $background.Dispose()
            $path.Dispose()
        }

        $shadowBrush = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(60, 0, 0, 0))
        $handleBrush = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(246, 232, 174))
        $handleHighlightBrush = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(255, 248, 213))
        $broomDarkBrush = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(160, 104, 43))
        $broomBrush = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(225, 167, 73))
        $bandBrush = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(75, 112, 92))
        try {
            $graphics.FillEllipse($shadowBrush, 20, 52, 34, 6)

            $graphics.TranslateTransform(32, 32)
            $graphics.RotateTransform(-35)
            $graphics.FillRectangle($handleBrush, (New-Object System.Drawing.Rectangle -ArgumentList -4, -26, 8, 39))
            $graphics.FillRectangle($handleHighlightBrush, (New-Object System.Drawing.Rectangle -ArgumentList -1, -25, 2, 37))
            $graphics.FillRectangle($bandBrush, (New-Object System.Drawing.Rectangle -ArgumentList -8, 10, 16, 7))
            $graphics.FillPolygon($broomDarkBrush, @(
                (New-Object System.Drawing.Point -ArgumentList -15, 17),
                (New-Object System.Drawing.Point -ArgumentList 15, 17),
                (New-Object System.Drawing.Point -ArgumentList 22, 36),
                (New-Object System.Drawing.Point -ArgumentList -22, 36)
            ))
            $graphics.FillPolygon($broomBrush, @(
                (New-Object System.Drawing.Point -ArgumentList -12, 18),
                (New-Object System.Drawing.Point -ArgumentList 12, 18),
                (New-Object System.Drawing.Point -ArgumentList 18, 34),
                (New-Object System.Drawing.Point -ArgumentList -18, 34)
            ))
            $graphics.ResetTransform()
        }
        finally {
            $shadowBrush.Dispose()
            $handleBrush.Dispose()
            $handleHighlightBrush.Dispose()
            $broomDarkBrush.Dispose()
            $broomBrush.Dispose()
            $bandBrush.Dispose()
        }

        $linePen = New-Object System.Drawing.Pen -ArgumentList ([System.Drawing.Color]::FromArgb(210, 255, 242, 205)), 2
        try {
            $graphics.TranslateTransform(32, 32)
            $graphics.RotateTransform(-35)
            $graphics.DrawLine($linePen, -10, 21, -14, 33)
            $graphics.DrawLine($linePen, 0, 20, 0, 34)
            $graphics.DrawLine($linePen, 10, 21, 14, 33)
            $graphics.ResetTransform()
        }
        finally {
            $linePen.Dispose()
        }

        return [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Set-DadaTaskbarIdentity {
    try {
        [void][DadaWin32.TaskbarIdentity]::SetCurrentProcessExplicitAppUserModelID($script:AppUserModelId)
    }
    catch {
    }
}

function Get-AppTempDirectory {
    return (Join-Path ([System.IO.Path]::GetTempPath()) $script:AppTempFolderName)
}

function Clear-DirectoryTreeQuietly {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try { $_.Attributes = [System.IO.FileAttributes]::Normal } catch { }
            }
            [System.IO.Directory]::Delete($Path, $true)
            return
        }
        catch {
            if ($attempt -lt 3) {
                Start-Sleep -Milliseconds (250 * $attempt)
            }
        }
    }
}

function Clear-StaleAppTempFiles {
    $directories = @(
        (Get-AppTempDirectory),
        (Join-Path ([System.IO.Path]::GetTempPath()) "ForestInventoryCleaner")
    ) | Select-Object -Unique

    foreach ($directory in $directories) {
        try {
            if (-not (Test-Path -LiteralPath $directory)) { continue }
            $now = Get-Date
            $alwaysDeletePatterns = @("guide_run_settings_*.json")
            $stalePatterns = @(
                "guide_run_progress_*.json",
                "guide_run_progress_*.json.tmp",
                "guide_run_result_*.json",
                "guide_run_cancel_*.flag",
                "guide_checks_*.mdb",
                "guide_checks_*.accdb",
                "guide_checks_*.ldb",
                "guide_checks_*.laccdb",
                "ai_context_*.mdb",
                "ai_context_*.accdb",
                "ai_context_*.ldb",
                "ai_context_*.laccdb",
                "~audit_export_*"
            )

            foreach ($pattern in $alwaysDeletePatterns) {
                Get-ChildItem -LiteralPath $directory -Filter $pattern -File -ErrorAction SilentlyContinue | ForEach-Object {
                    try { [System.IO.File]::Delete($_.FullName) } catch { }
                }
            }

            foreach ($pattern in $stalePatterns) {
                Get-ChildItem -LiteralPath $directory -Filter $pattern -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    try {
                        if (($now - $_.LastWriteTime).TotalHours -ge 12) {
                            if ($_.PSIsContainer) {
                                Clear-DirectoryTreeQuietly -Path $_.FullName
                            }
                            else {
                                [System.IO.File]::Delete($_.FullName)
                            }
                        }
                    }
                    catch {
                    }
                }
            }
        }
        catch {
        }
    }
}

function Field-Text-Expression {
    param([string]$FieldName)
    return "(" + (Quote-Name $FieldName) + " & '')"
}

function Normalize-FieldName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    return $Name.ToLowerInvariant() -replace "[^a-z0-9]", ""
}

function Test-ExcludedCleaningField {
    param([string]$FieldName)

    $field = Normalize-FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($field)) { return $false }

    return (
        $field -match "peracreexpan" -or
        $field -match "expan.*peracre" -or
        $field -match "acreexpan" -or
        $field -eq "calcsiteindex" -or
        $field -eq "gmp"
    )
}

function Get-RequiredPlotStatusFieldDefinitions {
    return @(
        [pscustomobject]@{ DisplayName = "Elevation"; CandidateNames = @("Elevation", "PlotElevation") },
        [pscustomobject]@{ DisplayName = "Aspect"; CandidateNames = @("Aspect", "PlotAspect") },
        [pscustomobject]@{ DisplayName = "Slope percent"; CandidateNames = @("SlopePercent", "SlopePct", "SlopePercentage", "Slope") },
        [pscustomobject]@{ DisplayName = "Slope position"; CandidateNames = @("SlopePosition", "SlopePositionCode", "SlopePos", "SlopePostion") },
        [pscustomobject]@{ DisplayName = "UTM northing"; CandidateNames = @("UTMNorthing", "UtmNorthing", "Northing") },
        [pscustomobject]@{ DisplayName = "UTM easting"; CandidateNames = @("UTMEasting", "UtmEasting", "Easting") },
        [pscustomobject]@{ DisplayName = "UTM zone"; CandidateNames = @("UTMZone", "UtmZone", "UTMZoneCode", "Zone") },
        [pscustomobject]@{ DisplayName = "Measurement date"; CandidateNames = @("MeasurementDate", "MeasDate", "DateMeasured", "MeasurementDt") },
        [pscustomobject]@{ DisplayName = "Crew"; CandidateNames = @("Crew", "CrewCode", "CrewID", "CrewName") },
        [pscustomobject]@{ DisplayName = "Stand class"; CandidateNames = @("StandClass", "StandClassCode", "StandCls") },
        [pscustomobject]@{ DisplayName = "Stand age"; CandidateNames = @("StandAge", "StandAgeCode") },
        [pscustomobject]@{ DisplayName = "Stockability percent"; CandidateNames = @("StockabilityPercent", "StockabilityPct", "StockabilityPercentage", "StockPct") },
        [pscustomobject]@{ DisplayName = "Stockability factor"; CandidateNames = @("StockabilityFactor", "StockFactor") }
    )
}

function Get-RequiredPlotStatusFieldDefinitionForField {
    param([string]$FieldName)

    $normalizedField = Normalize-FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($normalizedField)) { return $null }

    foreach ($definition in @(Get-RequiredPlotStatusFieldDefinitions)) {
        foreach ($candidate in @($definition.CandidateNames)) {
            if ($normalizedField -eq (Normalize-FieldName $candidate)) {
                return $definition
            }
        }
    }

    return $null
}

function Test-RequiredPlotStatusFieldName {
    param([string]$FieldName)

    return ($null -ne (Get-RequiredPlotStatusFieldDefinitionForField -FieldName $FieldName))
}

function Get-RequiredPlotStatusFieldDisplayName {
    param([string]$FieldName)

    $definition = Get-RequiredPlotStatusFieldDefinitionForField -FieldName $FieldName
    if ($null -eq $definition) { return $FieldName }
    return [string]$definition.DisplayName
}

function Get-RequiredTreeCrownFieldDefinitions {
    return @(
        [pscustomobject]@{
            DisplayName = "Crown ratio"
            RuleName = "Required crown ratio"
            CandidateNames = @("CrownRatio", "CrownRatioCode", "CrownRatioID")
        },
        [pscustomobject]@{
            DisplayName = "Crown class"
            RuleName = "Required crown class"
            CandidateNames = @("CrownClass", "CrownClassCode", "CrownClassID", "CrownPosition", "CrownPositionCode")
        }
    )
}

function Get-RequiredTreeCrownFieldDefinitionForField {
    param([string]$FieldName)

    $normalizedField = Normalize-FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($normalizedField)) { return $null }

    foreach ($definition in @(Get-RequiredTreeCrownFieldDefinitions)) {
        foreach ($candidate in @($definition.CandidateNames)) {
            if ($normalizedField -eq (Normalize-FieldName $candidate)) {
                return $definition
            }
        }
    }

    return $null
}

function Test-RequiredTreeCrownFieldName {
    param([string]$FieldName)

    return ($null -ne (Get-RequiredTreeCrownFieldDefinitionForField -FieldName $FieldName))
}

function Get-RequiredTreeRadialIncrementFieldDefinitions {
    return @(
        [pscustomobject]@{
            DisplayName = "Radial increment"
            RuleName = "Required radial increment"
            CandidateNames = @("RadialIncrement", "RadialIncr", "RadialInc", "RadIncrement", "RadInc")
        }
    )
}

function Get-RequiredTreeRadialIncrementFieldDefinitionForField {
    param([string]$FieldName)

    $normalizedField = Normalize-FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($normalizedField)) { return $null }

    foreach ($definition in @(Get-RequiredTreeRadialIncrementFieldDefinitions)) {
        foreach ($candidate in @($definition.CandidateNames)) {
            if ($normalizedField -eq (Normalize-FieldName $candidate)) {
                return $definition
            }
        }
    }

    return $null
}

function Test-RequiredTreeRadialIncrementFieldName {
    param([string]$FieldName)

    return ($null -ne (Get-RequiredTreeRadialIncrementFieldDefinitionForField -FieldName $FieldName))
}

function Normalize-AppColumnReviewKey {
    param(
        [string]$TableName,
        [string]$FieldName
    )

    return ((Normalize-FieldName $TableName) + "|" + (Normalize-FieldName $FieldName))
}

function ConvertTo-AppColumnActiveBoolean {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) { return $false }
    if ($Value -is [bool]) { return [bool]$Value }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    if ($text -match "^(?i:true|yes|y|on)$") { return $true }
    if ($text -match "^(?i:false|no|n|off)$") { return $false }

    $number = 0
    if ([int]::TryParse($text, [ref]$number)) {
        return ($number -ne 0)
    }

    return $false
}

function Get-AppColumnReviewMetadata {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    if ($null -eq $script:AppColumnReviewMetadataCache) {
        $script:AppColumnReviewMetadataCache = @{}
    }

    $sourceKey = ""
    try { $sourceKey = [string]$Connection.DataSource } catch { $sourceKey = "" }
    if ([string]::IsNullOrWhiteSpace($sourceKey)) { $sourceKey = "__current__" }
    if ($script:AppColumnReviewMetadataCache.ContainsKey($sourceKey)) {
        return $script:AppColumnReviewMetadataCache[$sourceKey]
    }

    $empty = [pscustomobject]@{
        HasMetadata = $false
        HasActiveColumn = $false
        ActiveMap = @{}
        Rows = @()
    }

    try {
        $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
        if (-not (Test-TableAvailable -Tables $tables -TableName "AppTables") -or
            -not (Test-TableAvailable -Tables $tables -TableName "AppColumns")) {
            $script:AppColumnReviewMetadataCache[$sourceKey] = $empty
            return $empty
        }

        $appTableColumns = @(Get-TableColumns -Connection $Connection -TableName "AppTables")
        $appColumnColumns = @(Get-TableColumns -Connection $Connection -TableName "AppColumns")
        if (-not (Test-ColumnExists -Columns $appTableColumns -Name "ID") -or
            -not (Test-ColumnExists -Columns $appTableColumns -Name "TableName") -or
            -not (Test-ColumnExists -Columns $appColumnColumns -Name "TableID") -or
            -not (Test-ColumnExists -Columns $appColumnColumns -Name "ColumnName")) {
            $script:AppColumnReviewMetadataCache[$sourceKey] = $empty
            return $empty
        }

        $hasActiveColumn = Test-ColumnExists -Columns $appColumnColumns -Name "Active"
        $hasCategoryType = Test-ColumnExists -Columns $appColumnColumns -Name "CategoryTypeID"
        $hasColumnId = Test-ColumnExists -Columns $appColumnColumns -Name "ID"
        $hasCodeTable = Test-TableAvailable -Tables $tables -TableName "AppColumnCodes"
        $canCountCodes = $false
        if ($hasCodeTable -and $hasColumnId) {
            $codeColumns = @(Get-TableColumns -Connection $Connection -TableName "AppColumnCodes")
            $canCountCodes =
                (Test-ColumnExists -Columns $codeColumns -Name "ColumnID") -and
                (Test-ColumnExists -Columns $codeColumns -Name "Code")
        }

        $columnIdSelect = if ($hasColumnId) { "c.[ID] AS [ColumnID]" } else { "'' AS [ColumnID]" }
        $categorySelect = if ($hasCategoryType) { "c.[CategoryTypeID] AS [CategoryTypeID]" } else { "0 AS [CategoryTypeID]" }
        $activeSelect = if ($hasActiveColumn) { "c.[Active] AS [Active]" } else { "1 AS [Active]" }
        $metadataSql = @"
SELECT t.[TableName], c.[ColumnName], $columnIdSelect, $categorySelect, $activeSelect
FROM [AppTables] AS t
INNER JOIN [AppColumns] AS c ON t.[ID] = c.[TableID]
ORDER BY t.[TableName], c.[ColumnName]
"@
        $metadata = Get-DataTable -Connection $Connection -Sql $metadataSql
        $codeCountMap = @{}
        if ($canCountCodes) {
            $codeRows = Get-DataTable -Connection $Connection -Sql "SELECT [ColumnID] FROM [AppColumnCodes] WHERE [Code] Is Not Null"
            foreach ($codeRow in $codeRows.Rows) {
                $codeColumnId = Get-DataRowText -Row $codeRow -ColumnName "ColumnID"
                if ([string]::IsNullOrWhiteSpace($codeColumnId)) { continue }
                if (-not $codeCountMap.ContainsKey($codeColumnId)) {
                    $codeCountMap[$codeColumnId] = 0
                }
                $codeCountMap[$codeColumnId] = [int]$codeCountMap[$codeColumnId] + 1
            }
        }

        $activeMap = @{}
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($row in $metadata.Rows) {
            $tableName = Get-DataRowText -Row $row -ColumnName "TableName"
            $columnName = Get-DataRowText -Row $row -ColumnName "ColumnName"
            if ([string]::IsNullOrWhiteSpace($tableName) -or [string]::IsNullOrWhiteSpace($columnName)) { continue }
            if (Test-ExcludedCleaningField -FieldName $columnName) { continue }

            $categoryTypeId = 0
            [void][int]::TryParse((Get-DataRowText -Row $row -ColumnName "CategoryTypeID"), [ref]$categoryTypeId)
            $codeCount = 0
            $columnId = Get-DataRowText -Row $row -ColumnName "ColumnID"
            if (-not [string]::IsNullOrWhiteSpace($columnId) -and $codeCountMap.ContainsKey($columnId)) {
                $codeCount = [int]$codeCountMap[$columnId]
            }
            $isActive = ConvertTo-AppColumnActiveBoolean -Value $row["Active"]
            $key = Normalize-AppColumnReviewKey -TableName $tableName -FieldName $columnName

            $activeMap[$key] = $isActive
            [void]$rows.Add([pscustomobject]@{
                TableName = $tableName
                FieldName = $columnName
                CategoryTypeId = $categoryTypeId
                CodeCount = $codeCount
                IsActive = $isActive
            })
        }

        $result = [pscustomobject]@{
            HasMetadata = $true
            HasActiveColumn = $hasActiveColumn
            ActiveMap = $activeMap
            Rows = @($rows.ToArray())
        }
        $script:AppColumnReviewMetadataCache[$sourceKey] = $result
        return $result
    }
    catch {
        $script:AppColumnReviewMetadataCache[$sourceKey] = $empty
        return $empty
    }
}

function Test-AppColumnFieldActiveForReview {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName
    )

    if (Test-ExcludedCleaningField -FieldName $FieldName) { return $false }

    $metadata = Get-AppColumnReviewMetadata -Connection $Connection
    if (-not [bool]$metadata.HasMetadata) {
        return $true
    }

    $key = Normalize-AppColumnReviewKey -TableName $TableName -FieldName $FieldName
    if ($metadata.ActiveMap.ContainsKey($key)) {
        return [bool]$metadata.ActiveMap[$key]
    }

    return $false
}

function Test-AppColumnFieldsActiveForReview {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string[]]$FieldNames
    )

    foreach ($fieldName in $FieldNames) {
        if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName $fieldName)) {
            return $false
        }
    }

    return $true
}

function Test-AnyRegenMinorPlotRuleEntered {
    foreach ($rule in @(Get-RegenMinorPlotRuleDefinitions)) {
        if (@($rule.MinorPlots).Count -gt 0) { return $true }
    }

    return $false
}

function Test-SupportedDataEntryReviewField {
    param(
        [string]$TableName,
        [string]$FieldName,
        [int]$CategoryTypeId = 0,
        [int]$CodeCount = 0
    )

    if (Test-ExcludedCleaningField -FieldName $FieldName) { return $false }
    if (Test-DmuMlField -FieldName $FieldName) { return $true }
    if ($CategoryTypeId -eq 3 -and $CodeCount -gt 0) { return $true }

    $table = Normalize-FieldName $TableName
    $field = Normalize-FieldName $FieldName
    if ($table -eq "treemeasurements" -and $field -in @("idbh", "totalheight", "stemcount", "treehistory", "problem1", "severity1", "problem2", "severity2", "treeclass", "treeclasscode", "treeclassid", "treeclasscd", "treeclasscdid", "treestatus", "treestatuscode", "treestatusid", "treestatuscd", "treestatuscdid")) { return $true }
    if ($table -eq "treemeasurements" -and (Test-RequiredTreeCrownFieldName -FieldName $FieldName)) { return $true }
    if ($table -eq "treemeasurements" -and (Test-RequiredTreeRadialIncrementFieldName -FieldName $FieldName)) { return $true }
    if ($table -eq "treemeasurements" -and $field -in @("speciescode", "species")) { return $true }
    if ($table -eq "trees" -and $field -in @("speciescode", "species", "treeclass", "treeclasscode", "treeclassid", "treeclasscd", "treeclasscdid", "treestatus", "treestatuscode", "treestatusid", "treestatuscd", "treestatuscdid")) { return $true }
    if ($table -eq "regenmeasurements" -and $field -in @("idbh", "stemcount", "speciescode", "species")) { return $true }
    if ($table -eq "regenmeasurements" -and $field -eq "minorplot") { return $true }
    if ($table -in @("plots", "plotmeasurements", "plotcustommeasurements") -and (Test-RequiredPlotStatusFieldName -FieldName $FieldName)) { return $true }
    if ($table -in @("plots", "plotmeasurements") -and $field -in @("plotremarks", "remarks", "remark", "plotnotes", "plotnote")) { return $true }
    if ($field -in @("idbh", "totalheight")) { return $true }

    return $false
}

function Get-ActiveReviewFieldsForExport {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $metadata = Get-AppColumnReviewMetadata -Connection $Connection
    if (-not [bool]$metadata.HasMetadata) { return @() }

    $targetTables = @(Get-CfiWorkbookTargetTables -Connection $Connection)
    $columnCache = @{}
    $reviewFields = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($metadata.Rows)) {
        if (-not [bool]$item.IsActive) { continue }
        $tableName = [string]$item.TableName
        $fieldName = [string]$item.FieldName
        if (-not ($targetTables -contains $tableName)) { continue }

        if (-not $columnCache.ContainsKey($tableName)) {
            try {
                $columnCache[$tableName] = @(Get-TableColumns -Connection $Connection -TableName $tableName)
            }
            catch {
                $columnCache[$tableName] = @()
            }
        }
        if (-not (Test-ColumnExists -Columns ([object[]]$columnCache[$tableName]) -Name $fieldName)) { continue }
        if (-not (Test-SupportedDataEntryReviewField -TableName $tableName -FieldName $fieldName -CategoryTypeId ([int]$item.CategoryTypeId) -CodeCount ([int]$item.CodeCount))) { continue }

        [void]$reviewFields.Add([pscustomobject]@{
            TableName = $tableName
            FieldName = $fieldName
            RuleName = "Reviewed - no data entry errors found"
            Note = "Data for $fieldName was reviewed and no data entry errors were found."
        })
    }

    if ($targetTables -contains "TreeMeasurements") {
        foreach ($pair in @(
            @{ Problem = "Problem1"; Severity = "Severity1" },
            @{ Problem = "Problem2"; Severity = "Severity2" }
        )) {
            $problemField = [string]$pair.Problem
            $severityField = [string]$pair.Severity
            if (-not (Test-AppColumnFieldsActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldNames @($problemField, $severityField))) { continue }
            if (-not $columnCache.ContainsKey("TreeMeasurements")) {
                try { $columnCache["TreeMeasurements"] = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements") } catch { $columnCache["TreeMeasurements"] = @() }
            }
            if (-not (Test-ColumnExists -Columns ([object[]]$columnCache["TreeMeasurements"]) -Name $problemField) -or
                -not (Test-ColumnExists -Columns ([object[]]$columnCache["TreeMeasurements"]) -Name $severityField)) {
                continue
            }

            $pairName = "$problemField/$severityField"
            [void]$reviewFields.Add([pscustomobject]@{
                TableName = "TreeMeasurements"
                FieldName = $pairName
                RuleName = "Reviewed - no data entry errors found"
                Note = "Data for $pairName was reviewed and no data entry errors were found."
            })
        }
    }

    return @($reviewFields.ToArray())
}

function Get-FieldManualTipsBySheet {
    if ($null -ne $script:FieldManualTipsBySheet) {
        return $script:FieldManualTipsBySheet
    }

    $script:FieldManualTipsBySheet = @{}
    $manualPath = Join-Path $PSScriptRoot "field-manual-text.json"
    if (-not (Test-Path -LiteralPath $manualPath)) {
        return $script:FieldManualTipsBySheet
    }

    try {
        $records = Get-Content -LiteralPath $manualPath -Raw | ConvertFrom-Json
        foreach ($record in $records) {
            $sheet = [string]$record.sheet
            $tip = $null
            foreach ($textRecord in @($record.texts)) {
                $text = [string]$textRecord.text
                if ([string]::IsNullOrWhiteSpace($text)) { continue }
                if ($text -eq "Show" -or $text -eq "Hide") { continue }
                $tip = $text
                break
            }

            if (-not [string]::IsNullOrWhiteSpace($sheet) -and -not [string]::IsNullOrWhiteSpace($tip)) {
                $script:FieldManualTipsBySheet[$sheet] = $tip
            }
        }
    }
    catch {
        $script:FieldManualTipsBySheet = @{}
    }

    return $script:FieldManualTipsBySheet
}

function Get-FieldManualSheetName {
    param(
        [string]$TableName,
        [string]$FieldName
    )

    $table = Normalize-FieldName $TableName
    $field = Normalize-FieldName $FieldName

    if ($field -eq "plotstatus") { return "Plot Status" }
    if ($field -eq "speciescode" -and $table -eq "regenmeasurements") { return "REGEN Species" }
    if ($field -eq "speciescode" -or $field -eq "sispecies" -or $field -eq "si2species" -or $field -eq "speciesprimary" -or $field -eq "speciessecondary") { return "Species" }
    if ($field -eq "treehistory" -or $field -eq "treestatus") { return "Tree History" }
    if (Test-RequiredTreeRadialIncrementFieldName -FieldName $FieldName) { return "Radial Increment" }
    if ($field -eq "idbh" -and $table -eq "regenmeasurements") { return "REGEN IDBH" }
    if ($field -eq "idbh") { return "IDBH" }
    if ($field -eq "totalheight" -or $field -eq "sitotalheight" -or $field -eq "si2totalheight" -or $field -eq "totalheightestimated") { return "Total Height" }
    if ($field -eq "minorplot" -and $table -eq "regenmeasurements") { return "REGEN Minor Plot" }
    if ($field -eq "minorplot") { return "Minor Plot" }
    if ($field -eq "stemcount" -and $table -eq "regenmeasurements") { return "REGEN Stem Count" }
    if ($field -eq "stemcount" -or $field -eq "livestemcount" -or $field -eq "totalstemcount") { return "StemCountWoodland" }
    if ($field -eq "crownratio" -or $field -eq "crownratiocode" -or $field -eq "crownratioid") { return "Crown Ratio" }
    if ($field -eq "crownclass" -or $field -eq "crownclasscode" -or $field -eq "crownclassid" -or $field -eq "crownposition" -or $field -eq "crownpositioncode") { return "Crown Class" }
    if ($field -eq "treeage" -or $field -eq "treeagecode") { return "Age Class" }
    if ($field -eq "bdftdefect" -or $field -eq "cuftdefect" -or $field -eq "totalcuftdefect") { return "Defect" }
    if ($field -eq "problem1") { return "Problem 1" }
    if ($field -eq "severity1") { return "Severity 1" }
    if ($field -eq "problem2") { return "Problem 2" }
    if ($field -eq "severity2") { return "Severity 2" }
    if ($field -eq "snagclass" -or $field -eq "snagcavity") { return "Snag Class-Category" }
    if ($field -eq "dmr") { return "DMR" }
    if ($field -eq "measurementdate" -or $field -eq "measdate" -or $field -eq "datemeasured" -or $field -eq "measurementdt") { return "Meas Date" }
    if ($field -eq "elevation" -or $field -eq "plotelevation") { return "Elevation" }
    if ($field -eq "aspect" -or $field -eq "plotaspect") { return "Aspect" }
    if ($field -eq "slopepercent" -or $field -eq "slopepct" -or $field -eq "slopepercentage" -or $field -eq "slope") { return "Slope%" }
    if ($field -eq "slopeposition" -or $field -eq "slopepositioncode" -or $field -eq "slopepos" -or $field -eq "slopepostion") { return "Slope Position" }
    if ($field -eq "utmnorthing" -or $field -eq "northing") { return "UTM Northing" }
    if ($field -eq "utmeasting" -or $field -eq "easting") { return "UTM Easting" }
    if ($field -eq "utmzone" -or $field -eq "utmzonecode" -or $field -eq "zone") { return "UTM Zone" }
    if ($field -eq "crew" -or $field -eq "crewcode" -or $field -eq "crewid" -or $field -eq "crewname") { return "Crew" }
    if ($field -eq "standclass" -or $field -eq "standclasscode" -or $field -eq "standcls") { return "Stand Class" }
    if ($field -eq "standage" -or $field -eq "standagecode") { return "Stand Age" }
    if ($field -eq "siage" -or $field -eq "si2age") { return "SI Age" }
    if ($field -eq "stockabilitypct" -or $field -eq "stockabilitypercent" -or $field -eq "stockabilitypercentage" -or $field -eq "stockpct" -or $field -eq "stockabilityfactor" -or $field -eq "stockfactor") { return "Stockability %" }
    if ($field -eq "covertype" -or $field -eq "undercovertype") { return "Cover Type" }
    if ($field -eq "sidbh" -or $field -eq "si2dbh") { return "SI IDBH" }
    if ($field -eq "sigrowth" -or $field -eq "si2growth") { return "SI Growth" }

    return $null
}

function Get-FieldManualTip {
    param(
        [string]$TableName,
        [string]$FieldName
    )

    $sheetName = Get-FieldManualSheetName -TableName $TableName -FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($sheetName)) { return $null }

    $tips = Get-FieldManualTipsBySheet
    if ($tips.ContainsKey($sheetName)) {
        return [string]$tips[$sheetName]
    }

    return $null
}

function Add-FieldManualTipToMessage {
    param(
        [string]$Message,
        [string]$TableName,
        [string]$FieldName
    )

    $tip = Get-FieldManualTip -TableName $TableName -FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($tip)) { return $Message }
    return "$Message Field manual guidance: $tip"
}

function Limit-TextLength {
    param(
        [string]$Text,
        [int]$MaxLength
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    if ($Text.Length -le $MaxLength) { return $Text }
    return $Text.Substring(0, [Math]::Max(0, $MaxLength - 3)) + "..."
}

function ConvertTo-AiSafeText {
    param(
        [object]$Value,
        [int]$MaxLength = 0
    )

    if ($null -eq $Value -or $Value -is [System.DBNull]) { return "" }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    $builder = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $text.Length; $i++) {
        $ch = $text[$i]
        $code = [int][char]$ch

        if ([char]::IsHighSurrogate($ch)) {
            if (($i + 1) -lt $text.Length -and [char]::IsLowSurrogate($text[$i + 1])) {
                [void]$builder.Append($ch)
                $i++
                [void]$builder.Append($text[$i])
            }
            else {
                [void]$builder.Append(" ")
            }
            continue
        }

        if ([char]::IsLowSurrogate($ch)) {
            [void]$builder.Append(" ")
            continue
        }

        if (($code -lt 32) -and $ch -notin @("`r", "`n", "`t")) {
            [void]$builder.Append(" ")
            continue
        }

        [void]$builder.Append($ch)
    }

    $safe = [System.Net.WebUtility]::HtmlDecode($builder.ToString())
    $safe = [regex]::Replace($safe, "[ \t]+", " ")
    $safe = [regex]::Replace($safe, "(\r?\n\s*){4,}", "`r`n`r`n")
    $safe = $safe.Trim()
    if ($MaxLength -gt 0) {
        $safe = Limit-TextLength -Text $safe -MaxLength $MaxLength
    }
    return $safe
}

function ConvertTo-Utf8JsonBody {
    param([string]$Json)

    return ,([System.Text.Encoding]::UTF8.GetBytes([string]$Json))
}

function Test-AiBadRequestMessage {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) { return $false }
    return ($Message -match "\(400\)|Bad Request|Azure rejected this specific request|maximum context|content filter|content management")
}

function Normalize-ManualText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $clean = [System.Net.WebUtility]::HtmlDecode($Text)
    $clean = [regex]::Replace($clean, "[\x00-\x08\x0B\x0C\x0E-\x1F]", " ")
    $clean = [regex]::Replace($clean, "[ \t]+", " ")
    $clean = [regex]::Replace($clean, "(\r?\n\s*){3,}", "`r`n`r`n")
    return $clean.Trim()
}

function Get-DocxTextFromFile {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $builder = New-Object System.Text.StringBuilder
        $entries = @($archive.Entries | Where-Object { $_.FullName -match "^word/(document|header\d*|footer\d*)\.xml$" })
        foreach ($entry in $entries) {
            $stream = $entry.Open()
            try {
                $reader = New-Object System.IO.StreamReader($stream)
                try {
                    $xml = $reader.ReadToEnd()
                    $xml = $xml -replace "<w:tab[^>]*/>", " "
                    $xml = $xml -replace "</w:p>", "`r`n"
                    $xml = $xml -replace "</w:tr>", "`r`n"
                    $xml = [regex]::Replace($xml, "<[^>]+>", " ")
                    [void]$builder.AppendLine($xml)
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $stream.Dispose()
            }
        }

        return Normalize-ManualText -Text $builder.ToString()
    }
    finally {
        $archive.Dispose()
    }
}

function Get-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) { return "" }

    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Convert-OpenXmlTextToPlainText {
    param([string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml)) { return "" }
    $text = $Xml -replace "<[^>]+>", " "
    return (Normalize-ManualText -Text $text)
}

function Get-SharedStringsFromOpenXmlWorkbook {
    param([System.IO.Compression.ZipArchive]$Archive)

    $sharedXml = Get-ZipEntryText -Archive $Archive -EntryName "xl/sharedStrings.xml"
    $strings = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($sharedXml)) { return $strings }

    foreach ($siMatch in [regex]::Matches($sharedXml, "<si\b[^>]*>(.*?)</si>", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $siXml = $siMatch.Groups[1].Value
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($textMatch in [regex]::Matches($siXml, "<t\b[^>]*>(.*?)</t>", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
            [void]$parts.Add((Normalize-ManualText -Text $textMatch.Groups[1].Value))
        }
        $text = [string]::Join("", [string[]]$parts.ToArray()).Trim()
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            [void]$strings.Add($text)
        }
    }

    return $strings
}

function Get-OpenXmlWorkbookTextFromFile {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $builder = New-Object System.Text.StringBuilder
        $sharedStrings = Get-SharedStringsFromOpenXmlWorkbook -Archive $archive
        if ($sharedStrings.Count -gt 0) {
            [void]$builder.AppendLine("Workbook text:")
            foreach ($text in $sharedStrings) {
                [void]$builder.AppendLine($text)
            }
        }

        $workbookXml = Get-ZipEntryText -Archive $archive -EntryName "xl/workbook.xml"
        if (-not [string]::IsNullOrWhiteSpace($workbookXml)) {
            [void]$builder.AppendLine("Workbook sheets:")
            foreach ($sheetMatch in [regex]::Matches($workbookXml, "<sheet\b[^>]*name=""([^""]+)""", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                [void]$builder.AppendLine((Normalize-ManualText -Text $sheetMatch.Groups[1].Value))
            }
        }

        foreach ($entry in @($archive.Entries | Where-Object { $_.FullName -match "^xl/worksheets/sheet\d+\.xml$" } | Sort-Object FullName)) {
            $stream = $entry.Open()
            try {
                $reader = New-Object System.IO.StreamReader($stream)
                try {
                    $xml = $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $stream.Dispose()
            }

            $sheetText = Convert-OpenXmlTextToPlainText -Xml $xml
            if (-not [string]::IsNullOrWhiteSpace($sheetText)) {
                [void]$builder.AppendLine($sheetText)
            }
        }

        return Normalize-ManualText -Text $builder.ToString()
    }
    finally {
        $archive.Dispose()
    }
}

function Expand-FlateBytes {
    param([byte[]]$Bytes)

    foreach ($skipBytes in @(0, 2)) {
        if ($Bytes.Length -le $skipBytes) { continue }
        $input = New-Object System.IO.MemoryStream(, $Bytes)
        $output = New-Object System.IO.MemoryStream
        try {
            if ($skipBytes -gt 0) {
                [void]$input.Seek($skipBytes, [System.IO.SeekOrigin]::Begin)
            }

            $deflate = New-Object System.IO.Compression.DeflateStream($input, [System.IO.Compression.CompressionMode]::Decompress)
            try {
                $buffer = New-Object byte[] 8192
                while (($read = $deflate.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $output.Write($buffer, 0, $read)
                }
            }
            finally {
                $deflate.Dispose()
            }

            if ($output.Length -gt 0) {
                return $output.ToArray()
            }
        }
        catch {
        }
        finally {
            $input.Dispose()
            $output.Dispose()
        }
    }

    return $null
}

function Convert-PdfLiteralStringToText {
    param([string]$PdfString)

    if ([string]::IsNullOrEmpty($PdfString)) { return "" }
    if ($PdfString.StartsWith("(") -and $PdfString.EndsWith(")")) {
        $PdfString = $PdfString.Substring(1, $PdfString.Length - 2)
    }

    $builder = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $PdfString.Length; $i++) {
        $char = $PdfString[$i]
        if ($char -ne "\") {
            [void]$builder.Append($char)
            continue
        }

        $i++
        if ($i -ge $PdfString.Length) { break }
        $escaped = $PdfString[$i]
        switch ($escaped) {
            "n" { [void]$builder.Append("`n"); break }
            "r" { [void]$builder.Append("`r"); break }
            "t" { [void]$builder.Append("`t"); break }
            "b" { [void]$builder.Append([char]8); break }
            "f" { [void]$builder.Append([char]12); break }
            "(" { [void]$builder.Append("("); break }
            ")" { [void]$builder.Append(")"); break }
            "\" { [void]$builder.Append("\"); break }
            "`r" {
                if (($i + 1) -lt $PdfString.Length -and $PdfString[$i + 1] -eq "`n") { $i++ }
                break
            }
            "`n" { break }
            default {
                if ($escaped -match "[0-7]") {
                    $octal = [string]$escaped
                    for ($j = 0; $j -lt 2 -and ($i + 1) -lt $PdfString.Length -and $PdfString[$i + 1] -match "[0-7]"; $j++) {
                        $i++
                        $octal += [string]$PdfString[$i]
                    }
                    [void]$builder.Append([char]([Convert]::ToInt32($octal, 8)))
                }
                else {
                    [void]$builder.Append($escaped)
                }
                break
            }
        }
    }

    return Normalize-ManualText -Text $builder.ToString()
}

function Convert-PdfHexStringToText {
    param([string]$HexString)

    if ([string]::IsNullOrWhiteSpace($HexString)) { return "" }
    $hex = $HexString.Trim()
    if ($hex.StartsWith("<") -and $hex.EndsWith(">")) {
        $hex = $hex.Substring(1, $hex.Length - 2)
    }
    $hex = $hex -replace "[^0-9A-Fa-f]", ""
    if ($hex.Length -eq 0) { return "" }
    if (($hex.Length % 2) -eq 1) { $hex += "0" }

    $bytes = New-Object byte[] ($hex.Length / 2)
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        $bytes[$i] = [Convert]::ToByte($hex.Substring($i * 2, 2), 16)
    }

    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return Normalize-ManualText -Text ([System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2))
    }

    return Normalize-ManualText -Text ([System.Text.Encoding]::GetEncoding(1252).GetString($bytes))
}

function Get-PdfTextFromContentStream {
    param([string]$Content)

    if ([string]::IsNullOrWhiteSpace($Content)) { return "" }
    $builder = New-Object System.Text.StringBuilder

    foreach ($match in [regex]::Matches($Content, "\((?:\\.|[^\\)])*\)\s*(?:Tj|'|"")", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $literalMatch = [regex]::Match($match.Value, "\((?:\\.|[^\\)])*\)", [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($literalMatch.Success) {
            $text = Convert-PdfLiteralStringToText -PdfString $literalMatch.Value
            if (-not [string]::IsNullOrWhiteSpace($text)) { [void]$builder.AppendLine($text) }
        }
    }

    foreach ($match in [regex]::Matches($Content, "<[0-9A-Fa-f\s]+>\s*Tj", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $hexMatch = [regex]::Match($match.Value, "<[0-9A-Fa-f\s]+>", [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($hexMatch.Success) {
            $text = Convert-PdfHexStringToText -HexString $hexMatch.Value
            if (-not [string]::IsNullOrWhiteSpace($text)) { [void]$builder.AppendLine($text) }
        }
    }

    foreach ($arrayMatch in [regex]::Matches($Content, "\[(.*?)\]\s*TJ", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $parts = New-Object System.Collections.Generic.List[string]
        $arrayText = $arrayMatch.Groups[1].Value
        foreach ($literalMatch in [regex]::Matches($arrayText, "\((?:\\.|[^\\)])*\)", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
            $text = Convert-PdfLiteralStringToText -PdfString $literalMatch.Value
            if (-not [string]::IsNullOrWhiteSpace($text)) { [void]$parts.Add($text) }
        }
        foreach ($hexMatch in [regex]::Matches($arrayText, "<[0-9A-Fa-f\s]+>", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
            $text = Convert-PdfHexStringToText -HexString $hexMatch.Value
            if (-not [string]::IsNullOrWhiteSpace($text)) { [void]$parts.Add($text) }
        }
        if ($parts.Count -gt 0) {
            [void]$builder.AppendLine([string]::Join("", [string[]]$parts.ToArray()))
        }
    }

    return Normalize-ManualText -Text $builder.ToString()
}

function Get-PdfTextFromStreams {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $singleByteEncoding = [System.Text.Encoding]::GetEncoding(28591)
    $pdfText = $singleByteEncoding.GetString($bytes)
    $builder = New-Object System.Text.StringBuilder

    foreach ($match in [regex]::Matches($pdfText, "<<(?<dict>.*?)>>\s*stream\s*(?:\r\n|\n|\r)?(?<data>.*?)\s*endstream", [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $dict = $match.Groups["dict"].Value
        $streamText = $match.Groups["data"].Value
        $streamBytes = $singleByteEncoding.GetBytes($streamText)

        if ($dict -match "/FlateDecode") {
            $expanded = Expand-FlateBytes -Bytes $streamBytes
            if ($null -eq $expanded) { continue }
            $streamBytes = $expanded
        }
        elseif ($dict -match "/DCTDecode|/JPXDecode|/CCITTFaxDecode|/JBIG2Decode") {
            continue
        }

        $content = $singleByteEncoding.GetString($streamBytes)
        $text = Get-PdfTextFromContentStream -Content $content
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            [void]$builder.AppendLine($text)
        }
    }

    return Normalize-ManualText -Text $builder.ToString()
}

function Get-PdfTextFromWordCom {
    param([string]$Path)

    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        try { $word.AutomationSecurity = 3 } catch { }
        $document = $word.Documents.Open($Path, $false, $true, $false)
        return Normalize-ManualText -Text ([string]$document.Content.Text)
    }
    finally {
        if ($null -ne $document) {
            try { $document.Close($false) } catch { }
            Release-ComObjectQuietly $document
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch { }
            Release-ComObjectQuietly $word
        }
    }
}

function Get-PdfTextFromFile {
    param([string]$Path)

    $text = ""
    try {
        $text = Get-PdfTextFromStreams -Path $Path
    }
    catch {
        $text = ""
    }

    if (-not [string]::IsNullOrWhiteSpace($text) -and $text.Length -ge 40) {
        return $text
    }

    try {
        $wordText = Get-PdfTextFromWordCom -Path $Path
        if (-not [string]::IsNullOrWhiteSpace($wordText)) {
            return $wordText
        }
    }
    catch {
    }

    if (-not [string]::IsNullOrWhiteSpace($text)) {
        return $text
    }

    throw "The PDF was opened, but no readable text was found. If it is scanned or image-only, run OCR or save it as DOCX/TXT first."
}

function Release-ComObjectQuietly {
    param([object]$ComObject)

    if ($null -eq $ComObject) { return }
    try {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
    catch {
    }
}

function Get-ExcelWorkbookTextFromFile {
    param([string]$Path)

    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.AutomationSecurity = 3 } catch { }
        try { $excel.EnableEvents = $false } catch { }
        $workbook = $excel.Workbooks.Open($Path, 3, $true)
        $builder = New-Object System.Text.StringBuilder

        foreach ($worksheet in @($workbook.Worksheets)) {
            try {
                [void]$builder.AppendLine("Sheet: $($worksheet.Name)")
                $usedRange = $worksheet.UsedRange
                try {
                    $values = $usedRange.Value2
                    if ($values -is [System.Array]) {
                        $rowLower = $values.GetLowerBound(0)
                        $rowUpper = $values.GetUpperBound(0)
                        $colLower = $values.GetLowerBound(1)
                        $colUpper = $values.GetUpperBound(1)
                        for ($rowIndex = $rowLower; $rowIndex -le $rowUpper; $rowIndex++) {
                            $parts = New-Object System.Collections.Generic.List[string]
                            for ($colIndex = $colLower; $colIndex -le $colUpper; $colIndex++) {
                                $value = $values.GetValue($rowIndex, $colIndex)
                                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                                    [void]$parts.Add(([string]$value).Trim())
                                }
                            }
                            if ($parts.Count -gt 0) {
                                [void]$builder.AppendLine([string]::Join(" | ", [string[]]$parts.ToArray()))
                            }
                        }
                    }
                    elseif ($null -ne $values -and -not [string]::IsNullOrWhiteSpace([string]$values)) {
                        [void]$builder.AppendLine([string]$values)
                    }
                }
                finally {
                    Release-ComObjectQuietly $usedRange
                }
            }
            finally {
                Release-ComObjectQuietly $worksheet
            }
        }

        return Normalize-ManualText -Text $builder.ToString()
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
            Release-ComObjectQuietly $workbook
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
            Release-ComObjectQuietly $excel
        }
    }
}

function Get-ProjectManualTextFromFile {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        throw "Choose a project field manual file first."
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        ".docx" { return Get-DocxTextFromFile -Path $Path }
        ".txt" { return Normalize-ManualText -Text (Get-Content -LiteralPath $Path -Raw) }
        ".md" { return Normalize-ManualText -Text (Get-Content -LiteralPath $Path -Raw) }
        ".csv" { return Normalize-ManualText -Text (Get-Content -LiteralPath $Path -Raw) }
        ".json" { return Normalize-ManualText -Text (Get-Content -LiteralPath $Path -Raw) }
        ".pdf" { return Get-PdfTextFromFile -Path $Path }
        ".xlsx" { return Get-OpenXmlWorkbookTextFromFile -Path $Path }
        ".xlsm" { return Get-OpenXmlWorkbookTextFromFile -Path $Path }
        ".xls" { return Get-ExcelWorkbookTextFromFile -Path $Path }
        default { throw "Project field manual uploads support PDF, DOCX, XLSX, XLSM, XLS, TXT, MD, CSV, or JSON files." }
    }
}

function Set-ProjectManualFile {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $script:ProjectManualText = Get-ProjectManualTextFromFile -Path $Path
    if ([string]::IsNullOrWhiteSpace($script:ProjectManualText)) {
        throw "The project manual was opened, but no readable text was found."
    }

    $script:ProjectManualPath = $Path
    $script:ProjectManualLoadedFrom = $Path
    $script:ProjectManualDigestByKey = @{}
    $script:AiGuidanceCache = @{}
    try {
        if ($null -ne $projectManualBox) { $projectManualBox.Text = $Path }
    }
    catch {
    }
}

function Get-ProjectManualText {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        return [string]$settings.ProjectManualText
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ProjectManualPath)) {
        if ($script:ProjectManualLoadedFrom -ne $script:ProjectManualPath -or [string]::IsNullOrWhiteSpace($script:ProjectManualText)) {
            Set-ProjectManualFile -Path $script:ProjectManualPath
        }
    }

    return $script:ProjectManualText
}

function Get-ProjectManualContextForAi {
    param([int]$MaxLength = 12000)

    $text = Get-ProjectManualText
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $path = $script:ProjectManualPath
    if ($null -ne $script:GuideRunSettings) {
        $path = [string]$script:GuideRunSettings.ProjectManualPath
    }
    return "Project field manual uploaded by user: $path`r`n" + (Limit-TextLength -Text $text -MaxLength $MaxLength)
}

function Get-ProjectManualSnippet {
    param(
        [string]$TableName,
        [string]$FieldName,
        [string]$RuleName,
        [string]$Message,
        [int]$MaxLength = 6000,
        [switch]$Dense
    )

    $text = Get-ProjectManualText
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    $terms = New-Object System.Collections.Generic.List[string]
    foreach ($source in @($TableName, $FieldName, $RuleName, $Message)) {
        foreach ($word in @(([string]$source) -split "[^A-Za-z0-9]+")) {
            if ($word.Length -ge 3) { Add-UniqueCode -Codes $terms -Code $word.ToLowerInvariant() }
        }
    }

    switch -Regex ($FieldName) {
        "IDBH|DBH" {
            foreach ($term in @("idbh", "dbh", "diameter", "woodland", "timber")) { Add-UniqueCode -Codes $terms -Code $term }
        }
        "TotalHeight|Height" {
            foreach ($term in @("height", "totalheight", "broken", "top")) { Add-UniqueCode -Codes $terms -Code $term }
        }
        "StemCount" {
            foreach ($term in @("stem", "count", "regen", "regeneration")) { Add-UniqueCode -Codes $terms -Code $term }
        }
        "SpeciesCode|Species" {
            foreach ($term in @("species", "code")) { Add-UniqueCode -Codes $terms -Code $term }
        }
        "TreeHistory" {
            foreach ($term in @("treehistory", "history", "live", "mortality", "harvest", "ingrowth")) { Add-UniqueCode -Codes $terms -Code $term }
        }
        "Problem|Severity" {
            foreach ($term in @("problem", "severity", "damage")) { Add-UniqueCode -Codes $terms -Code $term }
        }
    }

    $matches = New-Object System.Collections.Generic.List[object]
    foreach ($line in @($text -split "(\r?\n)+")) {
        $trimmed = $line.Trim()
        if ($trimmed.Length -lt 8) { continue }
        $lower = $trimmed.ToLowerInvariant()
        $score = 0
        foreach ($term in $terms) {
            if ($lower.Contains($term)) { $score++ }
        }
        if ($score -gt 0) {
            [void]$matches.Add([pscustomobject]@{ Score = $score; Text = $trimmed })
        }
    }

    $builder = New-Object System.Text.StringBuilder
    $seenLines = @{}
    $maxMatches = if ($Dense) { 8 } else { 30 }
    $maxLineLength = if ($Dense) { 180 } else { 1200 }
    foreach ($match in @($matches | Sort-Object Score -Descending | Select-Object -First $maxMatches)) {
        $lineText = Limit-TextLength -Text ([string]$match.Text) -MaxLength $maxLineLength
        $lineKey = ($lineText.ToLowerInvariant() -replace "\s+", " ").Trim()
        if ([string]::IsNullOrWhiteSpace($lineKey) -or $seenLines.ContainsKey($lineKey)) { continue }
        $seenLines[$lineKey] = $true
        [void]$builder.AppendLine($lineText)
        if ($builder.Length -ge $MaxLength) { break }
    }

    if ($builder.Length -eq 0) {
        if ($Dense) {
            $fallback = New-Object System.Text.StringBuilder
            $seenFallbackLines = @{}
            foreach ($line in @($text -split "(\r?\n)+")) {
                $trimmed = $line.Trim()
                if ($trimmed.Length -lt 8) { continue }
                $lineText = Limit-TextLength -Text $trimmed -MaxLength 180
                $lineKey = ($lineText.ToLowerInvariant() -replace "\s+", " ").Trim()
                if ([string]::IsNullOrWhiteSpace($lineKey) -or $seenFallbackLines.ContainsKey($lineKey)) { continue }
                $seenFallbackLines[$lineKey] = $true
                [void]$fallback.AppendLine($lineText)
                if ($fallback.Length -ge $MaxLength) { break }
            }
            return Limit-TextLength -Text $fallback.ToString() -MaxLength $MaxLength
        }

        return Limit-TextLength -Text $text -MaxLength $MaxLength
    }

    return Limit-TextLength -Text $builder.ToString() -MaxLength $MaxLength
}

function Get-ProjectManualDigestKey {
    param(
        [string]$TableName,
        [string]$FieldName,
        [string]$RuleName,
        [string]$Message
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($part in @($TableName, $FieldName, $RuleName, $Message)) {
        $clean = ([string]$part).ToLowerInvariant() -replace "[^a-z0-9]+", " "
        $clean = [regex]::Replace($clean.Trim(), "\s+", " ")
        if ($clean.Length -gt 80) { $clean = $clean.Substring(0, 80) }
        [void]$parts.Add($clean)
    }
    return [string]::Join("|", [string[]]$parts.ToArray())
}

function Convert-ManualSnippetToDigest {
    param(
        [string]$Snippet,
        [int]$MaxLength = 260
    )

    if ([string]::IsNullOrWhiteSpace($Snippet)) { return "" }

    $seen = @{}
    $phrases = New-Object System.Collections.Generic.List[string]
    foreach ($piece in [regex]::Split($Snippet, "(\r?\n)+|(?<=[\.;:])\s+")) {
        $text = ([string]$piece).Trim()
        if ($text.Length -lt 8) { continue }
        $text = [regex]::Replace($text, "[\x00-\x08\x0B\x0C\x0E-\x1F]", " ")
        $text = [regex]::Replace($text, "\s+", " ").Trim()
        $text = $text.Trim([char[]]@('-', '*', '"', [char]39, ' '))
        if ($text.Length -lt 8) { continue }
        $text = Limit-TextLength -Text $text -MaxLength 115
        $key = $text.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        [void]$phrases.Add($text.TrimEnd(".", ";", ":"))
        if ($phrases.Count -ge 3) { break }
    }

    if ($phrases.Count -eq 0) {
        return Limit-TextLength -Text ([regex]::Replace($Snippet.Trim(), "\s+", " ")) -MaxLength $MaxLength
    }

    return Limit-TextLength -Text ([string]::Join(" | ", [string[]]$phrases.ToArray())) -MaxLength $MaxLength
}

function Get-ProjectManualDigestForFinding {
    param(
        [string]$TableName,
        [string]$FieldName,
        [string]$RuleName,
        [string]$Message,
        [int]$MaxLength = 260
    )

    if ([string]::IsNullOrWhiteSpace((Get-ProjectManualText))) { return "" }
    if ($null -eq $script:ProjectManualDigestByKey) { $script:ProjectManualDigestByKey = @{} }

    $key = Get-ProjectManualDigestKey -TableName $TableName -FieldName $FieldName -RuleName $RuleName -Message $Message
    if ($script:ProjectManualDigestByKey.ContainsKey($key)) {
        return Limit-TextLength -Text ([string]$script:ProjectManualDigestByKey[$key]) -MaxLength $MaxLength
    }

    $snippet = Get-ProjectManualSnippet `
        -TableName $TableName `
        -FieldName $FieldName `
        -RuleName $RuleName `
        -Message $Message `
        -MaxLength 900 `
        -Dense

    $digest = Convert-ManualSnippetToDigest -Snippet $snippet -MaxLength 260
    $script:ProjectManualDigestByKey[$key] = $digest
    return Limit-TextLength -Text $digest -MaxLength $MaxLength
}

function Get-SelectedField {
    param([System.Windows.Forms.ComboBox]$Combo)
    if ($Combo.SelectedIndex -le 0) { return $null }
    return [string]$Combo.SelectedItem
}

function Get-GuideRunSettingText {
    param(
        [object]$Settings,
        [string]$Name
    )

    if ($null -eq $Settings -or [string]::IsNullOrWhiteSpace($Name)) { return "" }
    try {
        if ($Settings.PSObject.Properties[$Name]) {
            $value = $Settings.$Name
            if ($null -ne $value) { return [string]$value }
        }
    }
    catch {
    }

    return ""
}

function Get-GuideRunSettingBool {
    param(
        [object]$Settings,
        [string]$Name
    )

    $text = Get-GuideRunSettingText -Settings $Settings -Name $Name
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    try { return [bool]::Parse($text) } catch { return ([string]$text -match "^(1|yes|true)$") }
}

function Get-GuideRunLogPathFromExportPath {
    param([string]$ExportPath)

    $fallbackDirectory = Get-AppTempDirectory
    if ([string]::IsNullOrWhiteSpace($ExportPath)) {
        return (Join-Path $fallbackDirectory ("DADA_RunLog_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt"))
    }

    try {
        $directory = [System.IO.Path]::GetDirectoryName($ExportPath)
        if ([string]::IsNullOrWhiteSpace($directory)) { $directory = $fallbackDirectory }
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExportPath)
        if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = "DADA_RunLog_" + (Get-Date -Format "yyyyMMdd_HHmmss") }
        return (Join-Path $directory ($baseName + "_RunLog.txt"))
    }
    catch {
        return (Join-Path $fallbackDirectory ("DADA_RunLog_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt"))
    }
}

function ConvertTo-SafeExportFileProjectName {
    param([string]$Name)

    $text = ([string]$Name).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return "Project" }

    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $text = $text.Replace([string]$invalidChar, " ")
    }
    $text = $text -replace "\s+", "_"
    $text = $text -replace "_+", "_"
    $text = $text.Trim([char[]]"_. ")
    if ([string]::IsNullOrWhiteSpace($text)) { return "Project" }
    return (Limit-TextLength -Text $text -MaxLength 80)
}

function Get-DefaultExportProjectName {
    $fallbackName = "Project"
    if (-not [string]::IsNullOrWhiteSpace([string]$script:DbPath)) {
        try {
            $fallbackName = [System.IO.Path]::GetFileNameWithoutExtension([string]$script:DbPath)
        }
        catch {
        }
    }

    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$script:ConnectionString)) {
            $connection = Open-AccessConnection
            try {
                $tables = @(Get-UserTables -Connection $connection -IncludeAudit)
                if (Test-TableAvailable -Tables $tables -TableName "Projects") {
                    $columns = @(Get-TableColumns -Connection $connection -TableName "Projects")
                    foreach ($candidate in @("ProjectName", "Name", "Project", "Title", "Description", "ProjectDescription")) {
                        if (-not (Test-ColumnExists -Columns $columns -Name $candidate)) { continue }
                        $sql = "SELECT TOP 1 $(Quote-Name $candidate) FROM [Projects] WHERE $(Quote-Name $candidate) Is Not Null AND Trim(($(Quote-Name $candidate) & '')) <> '' AND InStr(1, ($(Quote-Name $candidate) & ''), 'Delete Me', 1) = 0"
                        $value = Get-Scalar -Connection $connection -Sql $sql
                        if ($null -ne $value -and -not ($value -is [System.DBNull]) -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                            return (ConvertTo-SafeExportFileProjectName -Name ([string]$value))
                        }
                    }
                }
            }
            finally {
                $connection.Close()
                $connection.Dispose()
            }
        }
    }
    catch {
    }

    return (ConvertTo-SafeExportFileProjectName -Name $fallbackName)
}

function Get-DefaultRunWorkbookFileName {
    $projectName = Get-DefaultExportProjectName
    return "${projectName}_CFIDataClean_" + (Get-Date -Format "yyyyMMdd_HHmm") + ".xlsx"
}

function Protect-GuideRunLogText {
    param(
        [string]$Text,
        [object]$Settings = $script:GuideRunSettings
    )

    if ($null -eq $Text) { return "" }
    $safeText = [string]$Text
    $secretValues = New-Object System.Collections.Generic.List[string]

    foreach ($name in @("AiApiKey", "Password")) {
        $value = Get-GuideRunSettingText -Settings $Settings -Name $name
        if (-not [string]::IsNullOrWhiteSpace($value)) { [void]$secretValues.Add($value) }
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$script:AiApiKey)) { [void]$secretValues.Add([string]$script:AiApiKey) }

    foreach ($secret in $secretValues) {
        if ([string]::IsNullOrWhiteSpace($secret)) { continue }
        $safeText = $safeText.Replace($secret, "[redacted]")
    }

    return $safeText
}

function Write-GuideRunLogEntry {
    param(
        [string]$Message,
        [object]$Settings = $script:GuideRunSettings
    )

    $logPath = [string]$script:GuideRunLogPath
    if ([string]::IsNullOrWhiteSpace($logPath)) {
        $logPath = Get-GuideRunSettingText -Settings $Settings -Name "RunLogPath"
    }
    if ([string]::IsNullOrWhiteSpace($logPath)) { return }

    try {
        $directory = [System.IO.Path]::GetDirectoryName($logPath)
        if (-not [string]::IsNullOrWhiteSpace($directory)) {
            [System.IO.Directory]::CreateDirectory($directory) | Out-Null
        }
        $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $safeMessage = Protect-GuideRunLogText -Text $Message -Settings $Settings
        $encoding = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::AppendAllText($logPath, "[$time] $safeMessage`r`n", $encoding)
    }
    catch {
    }
}

function Get-SqlPreviewForLog {
    param([string]$Sql)

    if ([string]::IsNullOrWhiteSpace($Sql)) { return "" }
    $preview = ([string]$Sql) -replace "\s+", " "
    return (Limit-TextLength -Text $preview.Trim() -MaxLength 1800)
}

function Initialize-GuideRunPerformance {
    $script:GuideRunPerformanceEntries = New-Object System.Collections.Generic.List[object]
    $script:GuideRunSlowSqlEntries = New-Object System.Collections.Generic.List[object]
    $script:GuideRunCurrentPerformanceStep = ""
    $script:GuideRunPerformanceSummaryWritten = $false
}

function Format-GuideRunDurationText {
    param([TimeSpan]$Elapsed)

    if ($Elapsed.TotalMinutes -ge 1) {
        return ("{0} ({1:N2}s)" -f (Format-ElapsedText -Elapsed $Elapsed), $Elapsed.TotalSeconds)
    }

    return ("{0:N2}s" -f $Elapsed.TotalSeconds)
}

function Format-ByteSizeText {
    param([long]$Bytes)

    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return "$Bytes bytes"
}

function Get-GuideRunResultCountText {
    param([object]$Result)

    if ($null -eq $Result) { return "" }
    if ($Result -is [System.Data.DataTable]) { return "$($Result.Rows.Count) row(s)" }
    if ($Result -is [int] -or $Result -is [long] -or $Result -is [decimal]) { return "$Result record(s)" }
    if ($Result -is [System.Array]) { return "$($Result.Count) item(s)" }
    return ""
}

function Add-GuideRunPerformanceEntry {
    param(
        [string]$Name,
        [TimeSpan]$Elapsed,
        [string]$Kind = "Step",
        [string]$CountText = "",
        [string]$Status = "completed",
        [string]$Detail = ""
    )

    if ($null -eq $script:GuideRunPerformanceEntries) {
        $script:GuideRunPerformanceEntries = New-Object System.Collections.Generic.List[object]
    }

    $durationText = Format-GuideRunDurationText -Elapsed $Elapsed
    $entry = [pscustomobject]@{
        Name = $Name
        Kind = $Kind
        Seconds = [Math]::Round($Elapsed.TotalSeconds, 3)
        Duration = $durationText
        Count = $CountText
        Status = $Status
        Detail = $Detail
        CompletedAt = (Get-Date).ToString("o")
    }
    [void]$script:GuideRunPerformanceEntries.Add($entry)

    $message = "Performance: $Kind '$Name' $Status in $durationText"
    if (-not [string]::IsNullOrWhiteSpace($CountText)) { $message += "; $CountText" }
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $message += "; $Detail" }
    Write-GuideRunLogEntry -Message $message
}

function Invoke-GuideTimedStep {
    param(
        [string]$Name,
        [scriptblock]$ScriptBlock,
        [string]$Kind = "Step"
    )

    $previousStep = [string]$script:GuideRunCurrentPerformanceStep
    $script:GuideRunCurrentPerformanceStep = $Name
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $result = $null
    try {
        $result = & $ScriptBlock
        $stopwatch.Stop()
        Add-GuideRunPerformanceEntry -Name $Name -Kind $Kind -Elapsed $stopwatch.Elapsed -CountText (Get-GuideRunResultCountText -Result $result)
        return $result
    }
    catch {
        $stopwatch.Stop()
        Add-GuideRunPerformanceEntry -Name $Name -Kind $Kind -Elapsed $stopwatch.Elapsed -Status "failed" -Detail (Limit-TextLength -Text $_.Exception.Message -MaxLength 220)
        throw
    }
    finally {
        $script:GuideRunCurrentPerformanceStep = $previousStep
    }
}

function Add-GuideRunSqlPerformanceEntry {
    param(
        [string]$Operation,
        [string]$Sql,
        [TimeSpan]$Elapsed,
        [string]$CountText = "",
        [switch]$Failed,
        [string]$ErrorMessage = ""
    )

    $isSlow = $Elapsed.TotalSeconds -ge [double]$script:GuideRunSlowSqlThresholdSeconds
    if (-not $Failed -and -not $isSlow) { return }

    if ($null -eq $script:GuideRunSlowSqlEntries) {
        $script:GuideRunSlowSqlEntries = New-Object System.Collections.Generic.List[object]
    }

    $preview = Get-SqlPreviewForLog -Sql $Sql
    $durationText = Format-GuideRunDurationText -Elapsed $Elapsed
    $entry = [pscustomobject]@{
        Operation = $Operation
        Step = [string]$script:GuideRunCurrentPerformanceStep
        Seconds = [Math]::Round($Elapsed.TotalSeconds, 3)
        Duration = $durationText
        Count = $CountText
        Failed = [bool]$Failed
        Error = $ErrorMessage
        SqlPreview = $preview
        CompletedAt = (Get-Date).ToString("o")
    }
    [void]$script:GuideRunSlowSqlEntries.Add($entry)

    $prefix = if ($Failed) { "SQL failed" } else { "Slow SQL" }
    $message = "${prefix}: $Operation in $durationText"
    if (-not [string]::IsNullOrWhiteSpace($CountText)) { $message += "; $CountText" }
    if (-not [string]::IsNullOrWhiteSpace([string]$script:GuideRunCurrentPerformanceStep)) { $message += "; step=$script:GuideRunCurrentPerformanceStep" }
    if ($Failed -and -not [string]::IsNullOrWhiteSpace($ErrorMessage)) { $message += "; error=$(Limit-TextLength -Text $ErrorMessage -MaxLength 180)" }
    if (-not [string]::IsNullOrWhiteSpace($preview)) { $message += "; SQL preview: $preview" }
    Write-GuideRunLogEntry -Message $message
}

function Get-GuideRunTopPerformanceEntries {
    param([int]$MaxCount = 8)

    if ($null -eq $script:GuideRunPerformanceEntries) { return @() }
    return @($script:GuideRunPerformanceEntries | Sort-Object Seconds -Descending | Select-Object -First $MaxCount)
}

function Write-GuideRunPerformanceSummary {
    param([string]$Status = "Completed")

    if ($script:GuideRunPerformanceSummaryWritten) { return }
    $script:GuideRunPerformanceSummaryWritten = $true

    Write-GuideRunLogEntry -Message ""
    Write-GuideRunLogEntry -Message "Performance summary: $Status"
    $topSteps = @(Get-GuideRunTopPerformanceEntries -MaxCount 12)
    if ($topSteps.Count -eq 0) {
        Write-GuideRunLogEntry -Message "Performance summary: no timed steps were recorded."
    }
    else {
        Write-GuideRunLogEntry -Message "Slowest timed steps:"
        $rank = 0
        foreach ($entry in $topSteps) {
            $rank++
            $line = "  $rank. $($entry.Name) - $($entry.Duration)"
            if (-not [string]::IsNullOrWhiteSpace([string]$entry.Count)) { $line += "; $($entry.Count)" }
            if ([string]$entry.Status -ne "completed") { $line += "; status=$($entry.Status)" }
            Write-GuideRunLogEntry -Message $line
        }
    }

    $slowSql = @()
    if ($null -ne $script:GuideRunSlowSqlEntries) {
        $slowSql = @($script:GuideRunSlowSqlEntries | Sort-Object Seconds -Descending | Select-Object -First 10)
    }
    if ($slowSql.Count -gt 0) {
        Write-GuideRunLogEntry -Message "Slowest Access SQL statements or failures:"
        $rank = 0
        foreach ($entry in $slowSql) {
            $rank++
            $line = "  $rank. $($entry.Operation) - $($entry.Duration)"
            if (-not [string]::IsNullOrWhiteSpace([string]$entry.Count)) { $line += "; $($entry.Count)" }
            if (-not [string]::IsNullOrWhiteSpace([string]$entry.Step)) { $line += "; step=$($entry.Step)" }
            if ([bool]$entry.Failed) { $line += "; failed=$($entry.Error)" }
            Write-GuideRunLogEntry -Message $line
            if (-not [string]::IsNullOrWhiteSpace([string]$entry.SqlPreview)) {
                Write-GuideRunLogEntry -Message ("     SQL preview: " + [string]$entry.SqlPreview)
            }
        }
    }
    else {
        Write-GuideRunLogEntry -Message "Slow SQL: none over $script:GuideRunSlowSqlThresholdSeconds second(s), and no SQL failures were recorded by the timing logger."
    }
}

function Write-GuideRunProgressLog {
    param(
        [string]$Status,
        [int]$Percent = -1
    )

    if ([string]::IsNullOrWhiteSpace($Status)) { return }
    if ([string]::IsNullOrWhiteSpace([string]$script:GuideRunLogPath)) { return }
    if ($Status -match "^(Preparing export row|Writing worksheet) ") { return }

    $entry = if ($Percent -ge 0) { "Progress $Percent%: $Status" } else { "Progress: $Status" }
    if ([string]$script:LastGuideRunLoggedProgress -eq $entry) { return }
    $script:LastGuideRunLoggedProgress = $entry
    Write-GuideRunLogEntry -Message $entry
}

function Initialize-GuideRunLog {
    param([object]$Settings)

    $logPath = Get-GuideRunSettingText -Settings $Settings -Name "RunLogPath"
    if ([string]::IsNullOrWhiteSpace($logPath)) { return }

    $script:GuideRunLogPath = $logPath
    $script:LastGuideRunLoggedProgress = ""
    try {
        $directory = [System.IO.Path]::GetDirectoryName($logPath)
        if (-not [string]::IsNullOrWhiteSpace($directory)) {
            [System.IO.Directory]::CreateDirectory($directory) | Out-Null
        }

        $startedAt = Get-GuideRunSettingText -Settings $Settings -Name "StartedAt"
        if ([string]::IsNullOrWhiteSpace($startedAt)) { $startedAt = (Get-Date).ToString("o") }
        $sourcePath = Get-GuideRunSettingText -Settings $Settings -Name "SourcePath"
        $exportPath = Get-GuideRunSettingText -Settings $Settings -Name "ExportPath"
        $manualPath = Get-GuideRunSettingText -Settings $Settings -Name "ProjectManualPath"
        $periodScope = if (Get-GuideRunSettingBool -Settings $Settings -Name "PeriodScopeEnabled") {
            if (Get-GuideRunSettingBool -Settings $Settings -Name "SingleMeasurementProject") {
                "Single measurement project; current period 1; TreeHistory checked across all periods."
            }
            else {
                "Current period $(Get-GuideRunSettingText -Settings $Settings -Name "CurrentPeriodNumber"); previous period $(Get-GuideRunSettingText -Settings $Settings -Name "PastPeriodNumber"); TreeHistory checked across all periods."
            }
        }
        else {
            "All periods included."
        }

        $builder = New-Object System.Text.StringBuilder
        [void]$builder.AppendLine("CFI DataClean run log")
        [void]$builder.AppendLine("App: $script:AppName")
        [void]$builder.AppendLine("Build: $script:AppVersion")
        [void]$builder.AppendLine("Started: $startedAt")
        [void]$builder.AppendLine("Source database: $sourcePath")
        [void]$builder.AppendLine("Export workbook: $exportPath")
        [void]$builder.AppendLine("Run log: $logPath")
        [void]$builder.AppendLine("Period scope: $periodScope")
        [void]$builder.AppendLine("AI guidance in export: $(if (Get-GuideRunSettingBool -Settings $Settings -Name "UseAiProjectGuidance") { "On" } else { "Off" })")
        [void]$builder.AppendLine("AI model/deployment: $(Get-GuideRunSettingText -Settings $Settings -Name "AiModel")")
        [void]$builder.AppendLine("Project manual: $manualPath")
        [void]$builder.AppendLine("")

        $safeText = Protect-GuideRunLogText -Text $builder.ToString() -Settings $Settings
        $encoding = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($logPath, $safeText, $encoding)
    }
    catch {
    }
}

function Write-GuideRunFailureLog {
    param(
        [string]$Status,
        [string]$Message,
        [string]$Detail = "",
        [object]$Settings = $script:GuideRunSettings
    )

    $elapsedText = Get-CurrentGuideRunElapsedText
    if ([string]::IsNullOrWhiteSpace($elapsedText)) { $elapsedText = "Not available" }

    Write-GuideRunLogEntry -Message ""
    Write-GuideRunLogEntry -Message "Run status: $Status" -Settings $Settings
    Write-GuideRunLogEntry -Message "Elapsed time: $elapsedText" -Settings $Settings
    Write-GuideRunLogEntry -Message "Error message: $Message" -Settings $Settings
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        Write-GuideRunLogEntry -Message "Error detail:" -Settings $Settings
        Write-GuideRunLogEntry -Message (Limit-TextLength -Text $Detail -MaxLength 12000) -Settings $Settings
    }
}

function Add-Log {
    param([string]$Message)

    try {
        if ($null -ne $form -and $form.InvokeRequired) {
            $messageCopy = $Message
            [void]$form.BeginInvoke([System.Action]{ Add-Log -Message $messageCopy })
            return
        }
    }
    catch {
    }

    $time = Get-Date -Format "HH:mm:ss"
    if ($null -ne $logBox) {
        $logBox.AppendText("[$time] $Message`r`n")
    }
    Write-GuideRunLogEntry -Message $Message
}

function Get-ElapsedStatusText {
    param(
        [string]$Status,
        [int]$Percent = -1,
        [bool]$Indeterminate = $false
    )

    if ([string]::IsNullOrWhiteSpace($Status)) { return $Status }
    if ($null -eq $script:GuideRunStartedAt) { return $Status }
    $elapsed = (Get-Date) - $script:GuideRunStartedAt
    $elapsedText = "$(Format-ElapsedText -Elapsed $elapsed) elapsed"
    return "$Status  $elapsedText"
}

function Write-GuideRunProgressFile {
    param(
        [string]$Status,
        [int]$Percent,
        [bool]$Indeterminate,
        [bool]$Clear
    )

    if ([string]::IsNullOrWhiteSpace($script:GuideRunProgressPath)) { return }

    try {
        $payload = [pscustomobject]@{
            Status = $Status
            Percent = $Percent
            Indeterminate = $Indeterminate
            Clear = $Clear
            UpdatedAt = (Get-Date).ToString("o")
        }
        $json = $payload | ConvertTo-Json -Depth 4
        $tempPath = "$script:GuideRunProgressPath.tmp"
        [System.IO.File]::WriteAllText($tempPath, $json, [System.Text.Encoding]::UTF8)
        if (Test-Path -LiteralPath $script:GuideRunProgressPath) {
            [System.IO.File]::Delete($script:GuideRunProgressPath)
        }
        [System.IO.File]::Move($tempPath, $script:GuideRunProgressPath)
    }
    catch {
    }
}

function Set-AppProgress {
    param(
        [string]$Status,
        [int]$Percent = -1,
        [switch]$Indeterminate,
        [switch]$Clear
    )

    try {
        Write-GuideRunProgressFile -Status $Status -Percent $Percent -Indeterminate ([bool]$Indeterminate) -Clear ([bool]$Clear)
        if ($env:FOREST_CLEANER_NO_UI -eq "1" -and -not [bool]$Clear) {
            Write-GuideRunProgressLog -Status $Status -Percent $Percent
        }

        if ($null -ne $form -and $form.InvokeRequired) {
            $statusCopy = $Status
            $percentCopy = $Percent
            $indeterminateCopy = [bool]$Indeterminate
            $clearCopy = [bool]$Clear
            [void]$form.BeginInvoke([System.Action]{ Set-AppProgress -Status $statusCopy -Percent $percentCopy -Indeterminate:$indeterminateCopy -Clear:$clearCopy })
            return
        }

        $displayStatus = if ($Clear) { "Ready" } else { Get-ElapsedStatusText -Status $Status -Percent $Percent -Indeterminate ([bool]$Indeterminate) }

        if ($null -ne $progressStatusLabel) {
            if ($Clear) {
                $progressStatusLabel.Text = "Ready"
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Status)) {
                $progressStatusLabel.Text = $displayStatus
            }
        }

        if ($null -ne $progressBar) {
            if ($Indeterminate) {
                $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
                $progressBar.MarqueeAnimationSpeed = 35
            }
            else {
                $progressBar.MarqueeAnimationSpeed = 0
                $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                if ($Clear) {
                    $progressBar.Value = 0
                }
                elseif ($Percent -ge 0) {
                    $safePercent = [Math]::Max(0, [Math]::Min(100, $Percent))
                    $progressBar.Value = $safePercent
                }
            }
        }

        if ($null -ne $statusStripLabel) {
            if ($Clear) {
                $statusStripLabel.Text = "Ready"
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Status)) {
                $statusStripLabel.Text = $displayStatus
            }
        }

        if ($null -ne $statusStripProgressBar) {
            if ($Indeterminate) {
                $statusStripProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
                $statusStripProgressBar.MarqueeAnimationSpeed = 35
            }
            else {
                $statusStripProgressBar.MarqueeAnimationSpeed = 0
                $statusStripProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                if ($Clear) {
                    $statusStripProgressBar.Value = 0
                }
                elseif ($Percent -ge 0) {
                    $safePercent = [Math]::Max(0, [Math]::Min(100, $Percent))
                    $statusStripProgressBar.Value = $safePercent
                }
            }
        }

        if ($null -ne $form) {
            if ($Clear) {
                $form.Text = $script:AppWindowBaseTitle
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Status)) {
                $titleStatus = Limit-TextLength -Text $displayStatus -MaxLength 70
                $form.Text = "$script:AppWindowBaseTitle - $titleStatus"
            }
        }

        [System.Windows.Forms.Application]::DoEvents()
    }
    catch {
    }
}

function Get-ProviderCandidates {
    param([string]$Path)

    $providers = @(
        "Microsoft.ACE.OLEDB.16.0",
        "Microsoft.ACE.OLEDB.15.0",
        "Microsoft.ACE.OLEDB.12.0"
    )

    if ([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -eq ".mdb") {
        $providers += "Microsoft.Jet.OLEDB.4.0"
    }

    return $providers
}

function New-AccessConnectionString {
    param(
        [string]$Path,
        [string]$Password
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "The selected database file does not exist."
    }

    $lastError = $null
    foreach ($provider in Get-ProviderCandidates $Path) {
        $candidate = "Provider=$provider;Data Source=$Path;Persist Security Info=False;"
        if (-not [string]::IsNullOrWhiteSpace($Password)) {
            $candidate += "Jet OLEDB:Database Password=$Password;"
        }

        $connection = New-Object System.Data.OleDb.OleDbConnection($candidate)
        try {
            $connection.Open()
            $connection.Close()
            return $candidate
        }
        catch {
            $lastError = $_.Exception.Message
        }
        finally {
            $connection.Dispose()
        }
    }

    $bitness = if ([Environment]::Is64BitProcess) { "64-bit" } else { "32-bit" }
    throw "Could not open the Access database from this $bitness process. Use the 32-bit launcher and make sure 32-bit Microsoft Access or the 32-bit Access Database Engine is installed. Last driver error: $lastError"
}

function Open-AccessConnection {
    if ([string]::IsNullOrWhiteSpace($script:ConnectionString)) {
        throw "Connect to a database first."
    }

    $connection = New-Object System.Data.OleDb.OleDbConnection($script:ConnectionString)
    $connection.Open()
    return $connection
}

function Invoke-NonQuery {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$Sql,
        [System.Data.OleDb.OleDbTransaction]$Transaction = $null
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = $Sql
    if ($null -ne $Transaction) {
        $command.Transaction = $Transaction
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $affected = $command.ExecuteNonQuery()
        $stopwatch.Stop()
        Add-GuideRunSqlPerformanceEntry -Operation "ExecuteNonQuery" -Sql $Sql -Elapsed $stopwatch.Elapsed -CountText "$affected affected row(s)"
        return $affected
    }
    catch {
        $stopwatch.Stop()
        Add-GuideRunSqlPerformanceEntry -Operation "ExecuteNonQuery" -Sql $Sql -Elapsed $stopwatch.Elapsed -Failed -ErrorMessage $_.Exception.Message
        $sqlPreview = Get-SqlPreviewForLog -Sql $Sql
        Write-GuideRunLogEntry -Message "Access write query failed. SQL preview: $sqlPreview"
        throw [System.InvalidOperationException]::new("Access write query failed: $($_.Exception.Message). SQL preview: $sqlPreview", $_.Exception)
    }
    finally {
        $command.Dispose()
    }
}

function Get-Scalar {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$Sql
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = $Sql
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $value = $command.ExecuteScalar()
        $stopwatch.Stop()
        Add-GuideRunSqlPerformanceEntry -Operation "ExecuteScalar" -Sql $Sql -Elapsed $stopwatch.Elapsed
        return $value
    }
    catch {
        $stopwatch.Stop()
        Add-GuideRunSqlPerformanceEntry -Operation "ExecuteScalar" -Sql $Sql -Elapsed $stopwatch.Elapsed -Failed -ErrorMessage $_.Exception.Message
        throw
    }
    finally {
        $command.Dispose()
    }
}

function Get-DataTable {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$Sql,
        [System.Data.OleDb.OleDbTransaction]$Transaction = $null
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = $Sql
    if ($null -ne $Transaction) {
        $command.Transaction = $Transaction
    }

    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
    $table = New-Object System.Data.DataTable
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        [void]$adapter.Fill($table)
        $stopwatch.Stop()
        Add-GuideRunSqlPerformanceEntry -Operation "Fill DataTable" -Sql $Sql -Elapsed $stopwatch.Elapsed -CountText "$($table.Rows.Count) row(s)"
        return ,$table
    }
    catch {
        $stopwatch.Stop()
        Add-GuideRunSqlPerformanceEntry -Operation "Fill DataTable" -Sql $Sql -Elapsed $stopwatch.Elapsed -Failed -ErrorMessage $_.Exception.Message
        throw
    }
    finally {
        $adapter.Dispose()
        $command.Dispose()
    }
}

function Get-AppInventoryTables {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string[]]$ExistingTables
    )

    if (-not ($ExistingTables -contains "AppTables")) { return @() }

    try {
        $table = Get-DataTable -Connection $Connection -Sql "SELECT [TableName] FROM [AppTables] ORDER BY [ID]"
        $names = New-Object System.Collections.Generic.List[string]
        foreach ($row in $table.Rows) {
            $name = [string]$row["TableName"]
            if (-not [string]::IsNullOrWhiteSpace($name) -and $ExistingTables -contains $name) {
                [void]$names.Add($name)
            }
        }

        return $names
    }
    catch {
        return @()
    }
}

function Get-UserTables {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [switch]$IncludeAudit
    )

    try {
        $schema = $Connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    }
    catch {
        $schema = $Connection.GetSchema("Tables")
    }

    $names = New-Object System.Collections.Generic.List[string]

    foreach ($row in $schema.Rows) {
        $name = [string]$row["TABLE_NAME"]
        $type = [string]$row["TABLE_TYPE"]
        if ($type -eq "TABLE" -and
            -not $name.StartsWith("MSys", [System.StringComparison]::OrdinalIgnoreCase) -and
            -not $name.StartsWith("~", [System.StringComparison]::OrdinalIgnoreCase) -and
            ($IncludeAudit -or -not $name.Equals("InventoryCleanAudit", [System.StringComparison]::OrdinalIgnoreCase))) {
            [void]$names.Add($name)
        }
    }

    $allNames = @($names)

    if (-not $IncludeAudit) {
        $appInventoryTables = @(Get-AppInventoryTables -Connection $Connection -ExistingTables $allNames)
        if ($appInventoryTables.Count -gt 0) {
            return $appInventoryTables
        }
    }

    return $allNames | Sort-Object
}

function Get-TableColumns {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $schema = $null
    try {
        $restrictions = New-Object "object[]" 4
        $restrictions[2] = $TableName
        $schema = $Connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Columns, $restrictions)
    }
    catch {
        $schema = $null
    }

    $columns = New-Object System.Collections.Generic.List[object]

    if ($null -ne $schema) {
        foreach ($row in $schema.Rows) {
            [void]$columns.Add([pscustomobject]@{
                Name = [string]$row["COLUMN_NAME"]
                Type = [int]$row["DATA_TYPE"]
                Ordinal = [int]$row["ORDINAL_POSITION"]
            })
        }
    }

    if ($columns.Count -eq 0) {
        $shape = Get-DataTable -Connection $Connection -Sql "SELECT TOP 0 * FROM $(Quote-Name $TableName)"
        foreach ($column in $shape.Columns) {
            [void]$columns.Add([pscustomobject]@{
                Name = [string]$column.ColumnName
                Type = Get-OleDbTypeFromDataColumn $column
                Ordinal = ([int]$column.Ordinal + 1)
            })
        }
    }

    return $columns | Sort-Object Ordinal
}

function Get-OleDbTypeFromDataColumn {
    param([System.Data.DataColumn]$Column)

    switch ($Column.DataType.FullName) {
        "System.String" { return 130 }
        "System.Int16" { return 2 }
        "System.Int32" { return 3 }
        "System.Single" { return 4 }
        "System.Double" { return 5 }
        "System.DateTime" { return 7 }
        "System.Boolean" { return 11 }
        "System.Byte" { return 17 }
        "System.Int64" { return 20 }
        "System.Decimal" { return 131 }
        default { return 0 }
    }
}

function Test-ColumnExists {
    param(
        [object[]]$Columns,
        [string]$Name
    )

    foreach ($column in $Columns) {
        if ($column.Name -eq $Name) { return $true }
    }
    return $false
}

function Test-TextColumn {
    param([int]$Type)
    return @(129, 130, 200, 201, 202, 203).Contains($Type)
}

function Test-NumericColumn {
    param([int]$Type)
    return @(2, 3, 4, 5, 6, 14, 16, 17, 18, 19, 20, 21, 131, 139).Contains($Type)
}

function Get-ColumnByName {
    param(
        [object[]]$Columns,
        [string]$Name
    )

    foreach ($column in $Columns) {
        if ($column.Name -eq $Name) { return $column }
    }

    return $null
}

function Find-CandidateColumn {
    param(
        [object[]]$Columns,
        [string[]]$Patterns,
        [switch]$NumericOnly,
        [switch]$TextOnly
    )

    foreach ($pattern in $Patterns) {
        foreach ($column in $Columns) {
            $normalized = $column.Name.ToLowerInvariant() -replace "[^a-z0-9]", ""
            if ($normalized -match $pattern) {
                if ($NumericOnly -and -not (Test-NumericColumn $column.Type)) { continue }
                if ($TextOnly -and -not (Test-TextColumn $column.Type)) { continue }
                return $column.Name
            }
        }
    }

    return $null
}

function Set-ComboItems {
    param(
        [System.Windows.Forms.ComboBox]$Combo,
        [object[]]$Columns,
        [string]$SelectedName
    )

    $Combo.Items.Clear()
    [void]$Combo.Items.Add("(none)")
    foreach ($column in $Columns) {
        [void]$Combo.Items.Add($column.Name)
    }

    if (-not [string]::IsNullOrWhiteSpace($SelectedName) -and $Combo.Items.Contains($SelectedName)) {
        $Combo.SelectedItem = $SelectedName
    }
    else {
        $Combo.SelectedIndex = 0
    }
}

function Ensure-AuditTable {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = Get-UserTables -Connection $Connection -IncludeAudit
    if ($tables -contains "InventoryCleanAudit") {
        Ensure-AuditExtraColumns -Connection $Connection
        return
    }

    $sql = @"
CREATE TABLE [InventoryCleanAudit] (
    [AuditId] AUTOINCREMENT PRIMARY KEY,
    [AuditTime] DATETIME,
    [TableName] TEXT(128),
    [RuleName] TEXT(128),
    [RecordLabel] TEXT(255),
    [FieldName] TEXT(128),
    [SourceRowId] TEXT(255),
    [PlotNumber] TEXT(255),
    [TreeNumber] TEXT(255),
    [MinorPlot] TEXT(255),
    [PeriodNumber] TEXT(255),
    [PlotKey] TEXT(255),
    [TreeKey] TEXT(255),
    [RegenMeasKey] TEXT(255),
    [SpeciesCode] TEXT(255),
    [PlotStatus] TEXT(255),
    [PlotRemarks] MEMO,
    [RegenRemarks] MEMO,
    [ObservedValue] TEXT(255),
    [Message] MEMO
)
"@
    [void](Invoke-NonQuery -Connection $Connection -Sql $sql)
    Ensure-AuditExtraColumns -Connection $Connection
}

function Ensure-AuditExtraColumns {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = Get-UserTables -Connection $Connection -IncludeAudit
    if (-not ($tables -contains "InventoryCleanAudit")) { return }

    $columns = @(Get-TableColumns -Connection $Connection -TableName "InventoryCleanAudit")
    foreach ($columnName in @("SourceRowId", "PlotNumber", "TreeNumber", "MinorPlot", "PeriodNumber", "PlotKey", "TreeKey", "RegenMeasKey", "SpeciesCode", "PlotStatus", "PlotRemarks", "RegenRemarks")) {
        if (-not (Test-ColumnExists -Columns $columns -Name $columnName)) {
            $columnType = if ($columnName -in @("PlotRemarks", "RegenRemarks")) { "MEMO" } else { "TEXT(255)" }
            [void](Invoke-NonQuery -Connection $Connection -Sql "ALTER TABLE [InventoryCleanAudit] ADD COLUMN $(Quote-Name $columnName) $columnType")
        }
    }
}

function Get-SourceRowIdExpression {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    foreach ($candidate in @("MeasurementID", "TreeID", "PlotID", "ProjectPeriodID", "ID")) {
        if (Test-ColumnExists -Columns $columns -Name $candidate) {
            return "(" + (Quote-Name $candidate) + " & '')"
        }
    }

    return "''"
}

function Ensure-NeedsReviewColumn {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $columns = Get-TableColumns -Connection $Connection -TableName $TableName
    if (Test-ColumnExists -Columns $columns -Name "NeedsReview") { return }

    $sql = "ALTER TABLE $(Quote-Name $TableName) ADD COLUMN [NeedsReview] YESNO"
    [void](Invoke-NonQuery -Connection $Connection -Sql $sql)
}

function Get-RecordLabelExpression {
    param(
        [string]$PlotField,
        [string]$TreeField
    )

    if ($PlotField -and $TreeField) {
        return (Field-Text-Expression $PlotField) + " & '/' & " + (Field-Text-Expression $TreeField)
    }

    if ($PlotField) { return Field-Text-Expression $PlotField }
    if ($TreeField) { return Field-Text-Expression $TreeField }
    return "'record'"
}

function Add-RangeAudit {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.OleDb.OleDbTransaction]$Transaction,
        [string]$TableName,
        [string]$FieldName,
        [decimal]$MaximumValue,
        [string]$RuleName,
        [string]$Message,
        [string]$RecordLabelExpression
    )

    $field = Quote-Name $FieldName
    $sourceRowId = Get-SourceRowIdExpression -Connection $Connection -TableName $TableName
    $maximumSql = $MaximumValue.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $condition = "$field Is Null OR $field <= 0 OR $field > $maximumSql"
    $condition = Add-PeriodScopeToCondition -Connection $Connection -TableName $TableName -RuleName $RuleName -FieldName $FieldName -Condition $condition
    $auditSql = @"
INSERT INTO [InventoryCleanAudit]
    ([AuditTime], [TableName], [RuleName], [RecordLabel], [FieldName], [SourceRowId], [ObservedValue], [Message])
SELECT
    Now(),
    $(Sql-Text $TableName),
    $(Sql-Text $RuleName),
    $RecordLabelExpression,
    $(Sql-Text $FieldName),
    $sourceRowId,
    ($field & ''),
    $(Sql-Text $Message)
FROM $(Quote-Name $TableName)
WHERE $condition
"@
    $rows = Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql $auditSql

    $markSql = "UPDATE $(Quote-Name $TableName) SET [NeedsReview] = True WHERE $condition"
    [void](Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql $markSql)

    return $rows
}

function Normalize-CodeValue {
    param([object]$Value)

    if ($null -eq $Value -or [System.DBNull]::Value.Equals($Value)) { return "" }

    $text = [string]$Value
    foreach ($character in @([char]0x200B, [char]0x200C, [char]0x200D, [char]0xFEFF)) {
        $text = $text.Replace([string]$character, "")
    }
    $text = $text.Replace([string][char]0x00A0, " ")
    return $text.Trim()
}

function Add-UniqueCode {
    param(
        [System.Collections.Generic.List[string]]$Codes,
        [string]$Code
    )

    if ([string]::IsNullOrWhiteSpace($Code)) { return }
    if (-not $Codes.Contains($Code)) {
        [void]$Codes.Add($Code)
    }
}

function Test-DmuMlField {
    param([string]$FieldName)

    return ([string]$FieldName).Trim().Equals("DMUML", [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-DmuMlAuditCondition {
    param([string]$QuotedField)

    $rawValue = "Trim(($QuotedField & ''))"
    $normalizedValue = "Right('000' & $rawValue, 3)"
    $sumExpression = "Val(Mid($normalizedValue, 1, 1)) + Val(Mid($normalizedValue, 2, 1)) + Val(Mid($normalizedValue, 3, 1))"

    return "$rawValue <> '' AND (Len($rawValue) > 3 OR $rawValue Like '*[!0-9]*' OR $normalizedValue Like '*[!0-3]*' OR ($sumExpression) > 6)"
}

function Get-DmuMlVerificationSql {
    param(
        [string]$TableName,
        [string]$FieldName
    )

    $table = Quote-Name $TableName
    $field = Quote-Name $FieldName
    $condition = Get-DmuMlAuditCondition -QuotedField $field
    return "SELECT * FROM $table WHERE $condition;"
}

function Get-CodeGreaterThanZeroCondition {
    param([string]$QuotedField)

    $textValue = "Trim(($QuotedField & ''))"
    return "($QuotedField Is Not Null AND $textValue <> '' AND Val($textValue) > 0)"
}

function Get-CodeNoneCondition {
    param([string]$QuotedField)

    $textValue = "Trim(($QuotedField & ''))"
    return "($QuotedField Is Null OR $textValue = '' OR Val($textValue) = 0)"
}

function Get-RegenSpeciesRecordedCondition {
    param([string]$SpeciesField = "[SpeciesCode]")

    return Get-CodeGreaterThanZeroCondition -QuotedField $SpeciesField
}

function Get-FieldHasRecordedValueCondition {
    param([string]$FieldExpression)

    $valueText = "Trim(($FieldExpression & ''))"
    return "($FieldExpression Is Not Null AND $valueText <> '')"
}

function Get-RegenSpeciesRequiredCondition {
    param(
        [object[]]$Columns,
        [string]$SpeciesField = "[SpeciesCode]"
    )

    $speciesValue = "Trim(($SpeciesField & ''))"
    $recordedConditions = New-Object System.Collections.Generic.List[string]
    foreach ($fieldName in @("RegenMeasKey", "IDBH", "StemCount", "MinorPlot")) {
        if (-not (Test-ColumnExists -Columns $Columns -Name $fieldName)) { continue }
        $field = Quote-Name $fieldName
        [void]$recordedConditions.Add((Get-FieldHasRecordedValueCondition -FieldExpression $field))
    }

    if ($recordedConditions.Count -eq 0) { return "" }

    $recordedCondition = [string]::Join(" OR ", [string[]]$recordedConditions.ToArray())
    return "(($SpeciesField Is Null OR $speciesValue = '') AND ($recordedCondition))"
}

function Get-RegenIdbhRequiredCondition {
    param(
        [string]$SpeciesField = "[SpeciesCode]",
        [string]$IdbhField = "[IDBH]"
    )

    $speciesRecorded = Get-RegenSpeciesRecordedCondition -SpeciesField $SpeciesField
    $idbhValue = "Trim(($IdbhField & ''))"
    return "(($speciesRecorded) AND ($IdbhField Is Null OR $idbhValue = '' OR $idbhValue Not In ('0', '20', '40')))"
}

function Get-RegenStemCountRequiredCondition {
    param(
        [string]$SpeciesField = "[SpeciesCode]",
        [string]$StemCountField = "[StemCount]",
        [decimal]$MaximumValue
    )

    $speciesRecorded = Get-RegenSpeciesRecordedCondition -SpeciesField $SpeciesField
    $stemValue = "Trim(($StemCountField & ''))"
    $maximumSql = $MaximumValue.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $stemRecorded = Get-FieldHasRecordedValueCondition -FieldExpression $StemCountField
    $enteredOutOfRange = "(($stemRecorded) AND (Val($stemValue) <= 0 OR Val($stemValue) > $maximumSql))"
    $missingForRecordedSpecies = "(($speciesRecorded) AND ($StemCountField Is Null OR $stemValue = ''))"
    return "(($enteredOutOfRange) OR ($missingForRecordedSpecies))"
}

function Get-RegenMinorPlotRequiredCondition {
    param(
        [string]$SpeciesField = "[SpeciesCode]",
        [string]$MinorPlotField = "[MinorPlot]"
    )

    $speciesRecorded = Get-RegenSpeciesRecordedCondition -SpeciesField $SpeciesField
    $minorPlotValue = "Trim(($MinorPlotField & ''))"
    return "(($speciesRecorded) AND ($MinorPlotField Is Null OR $minorPlotValue = ''))"
}

function Get-ProblemSeverityMismatchCondition {
    param(
        [string]$ProblemField,
        [string]$SeverityField
    )

    $problem = Quote-Name $ProblemField
    $severity = Quote-Name $SeverityField
    $problemEntered = Get-CodeGreaterThanZeroCondition -QuotedField $problem
    $severityEntered = Get-CodeGreaterThanZeroCondition -QuotedField $severity
    $problemNone = Get-CodeNoneCondition -QuotedField $problem
    $severityNone = Get-CodeNoneCondition -QuotedField $severity

    return "(($problemEntered AND $severityNone) OR ($severityEntered AND $problemNone))"
}

function Get-NewMortalityProblem1SeverityCondition {
    param(
        [string]$TreeHistoryField = "[TreeHistory]",
        [string]$ProblemField = "[Problem1]",
        [string]$SeverityField = "[Severity1]"
    )

    $historyValue = "Trim(($TreeHistoryField & ''))"
    $problemEntered = Get-CodeGreaterThanZeroCondition -QuotedField $ProblemField
    $severityValue = "Trim(($SeverityField & ''))"

    return "($TreeHistoryField Is Not Null AND $historyValue <> '' AND Val($historyValue) In (2, 3) AND ($problemEntered) AND ($SeverityField Is Null OR $severityValue = '' OR Val($severityValue) <> 3))"
}

function Add-AuditRow {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.OleDb.OleDbTransaction]$Transaction = $null,
        [string]$TableName,
        [string]$RuleName,
        [string]$RecordLabel,
        [string]$FieldName,
        [string]$ObservedValue,
        [string]$Message
    )

    $sql = @"
INSERT INTO [InventoryCleanAudit]
    ([AuditTime], [TableName], [RuleName], [RecordLabel], [FieldName], [ObservedValue], [Message])
VALUES
    (Now(), $(Sql-Text $TableName), $(Sql-Text $RuleName), $(Sql-Text $RecordLabel), $(Sql-Text $FieldName), $(Sql-Text $ObservedValue), $(Sql-Text $Message))
"@
    [void](Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql $sql)
}

function Add-ConditionAudit {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.OleDb.OleDbTransaction]$Transaction,
        [string]$TableName,
        [string]$RuleName,
        [string]$FieldName,
        [string]$ObservedExpression,
        [string]$Message,
        [string]$RecordLabelExpression,
        [string]$Condition
    )

    $observed = "Left((($ObservedExpression) & ''), 255)"
    $sourceRowId = Get-SourceRowIdExpression -Connection $Connection -TableName $TableName
    $Condition = Add-PeriodScopeToCondition -Connection $Connection -TableName $TableName -RuleName $RuleName -FieldName $FieldName -Condition $Condition
    $auditSql = @"
INSERT INTO [InventoryCleanAudit]
    ([AuditTime], [TableName], [RuleName], [RecordLabel], [FieldName], [SourceRowId], [ObservedValue], [Message])
SELECT
    Now(),
    $(Sql-Text $TableName),
    $(Sql-Text $RuleName),
    $RecordLabelExpression,
    $(Sql-Text $FieldName),
    $sourceRowId,
    $observed,
    $(Sql-Text $Message)
FROM $(Quote-Name $TableName)
WHERE $Condition
"@
    $rows = Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql $auditSql

    $markSql = "UPDATE $(Quote-Name $TableName) SET [NeedsReview] = True WHERE $Condition"
    [void](Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql $markSql)

    return $rows
}

function Add-CodeValidationAudit {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.OleDb.OleDbTransaction]$Transaction,
        [string]$TableName,
        [object[]]$Columns
    )

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not ($tables -contains "AppTables") -or
        -not ($tables -contains "AppColumns") -or
        -not ($tables -contains "AppColumnCodes")) {
        return 0
    }

    $metadataSql = @"
SELECT c.[ID], c.[ColumnName]
FROM [AppTables] AS t
INNER JOIN [AppColumns] AS c ON t.[ID] = c.[TableID]
WHERE t.[TableName] = $(Sql-Text $TableName)
  AND c.[CategoryTypeID] = 3
"@
    $metadata = Get-DataTable -Connection $Connection -Transaction $Transaction -Sql $metadataSql
    if ($metadata.Rows.Count -eq 0) { return 0 }

    $codesTable = Get-DataTable -Connection $Connection -Transaction $Transaction -Sql "SELECT [ColumnID], [Code], [CodeLabel] FROM [AppColumnCodes]"
    $plotField = Find-CandidateColumn -Columns $Columns -Patterns @("^plotmeaskey$", "^plotkey$", "^plotid$", "^plotnumber$", "^plot$", "plotid", "plotnumber", "stand", "unit")
    $treeField = Find-CandidateColumn -Columns $Columns -Patterns @("^treemeaskey$", "^regenmeaskey$", "^treekey$", "^treeid$", "^treenumber$", "^tree$", "tag", "stem")
    $recordLabel = Get-RecordLabelExpression -PlotField $plotField -TreeField $treeField

    $auditCount = 0
    foreach ($row in $metadata.Rows) {
        $columnName = [string]$row["ColumnName"]
        if (-not (Test-ColumnExists -Columns $Columns -Name $columnName)) { continue }
        if (Test-ExcludedCleaningField -FieldName $columnName) { continue }
        if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName $columnName)) { continue }

        if (Test-DmuMlField -FieldName $columnName) {
            $field = Quote-Name $columnName
            $condition = Get-DmuMlAuditCondition -QuotedField $field
            $message = Add-FieldManualTipToMessage `
                -Message "DMUML must be a three-position numeric code. Use digits 0, 1, 2, or 3 only, and the three digits must add up to no more than 6. Short numeric entries are read as left-padded codes, so 20 is checked as 020." `
                -TableName $TableName `
                -FieldName $columnName

            $auditCount += Add-ConditionAudit `
                -Connection $Connection `
                -Transaction $Transaction `
                -TableName $TableName `
                -RuleName "DMUML code check" `
                -FieldName $columnName `
                -ObservedExpression "Trim(($field & ''))" `
                -Message $message `
                -RecordLabelExpression $recordLabel `
                -Condition $condition
            continue
        }

        $columnId = [string]$row["ID"]
        $validCodes = New-Object System.Collections.Generic.List[string]
        $codeExamples = New-Object System.Collections.Generic.List[string]
        foreach ($codeRow in $codesTable.Rows) {
            if ([string]$codeRow["ColumnID"] -ne $columnId) { continue }

            $code = Normalize-CodeValue $codeRow["Code"]
            Add-UniqueCode -Codes $validCodes -Code $code
            $label = Normalize-CodeValue $codeRow["CodeLabel"]
            if (-not [string]::IsNullOrWhiteSpace($code) -and -not [string]::IsNullOrWhiteSpace($label)) {
                Add-UniqueCode -Codes $codeExamples -Code "$code = $label"
            }
            if ($code -match "^\d+$") {
                Add-UniqueCode -Codes $validCodes -Code (([int64]$code).ToString())
            }
        }

        if ($validCodes.Count -eq 0) { continue }

        $field = Quote-Name $columnName
        $validList = ($validCodes | Sort-Object -Unique | ForEach-Object { Sql-Text $_ }) -join ", "
        $normalizedField = "Trim(($field & ''))"
        $observedExpression = $normalizedField
        $condition = "$normalizedField <> '' AND $normalizedField Not In ($validList)"
        $message = Add-FieldManualTipToMessage `
            -Message "The value is not listed as a valid code in the CFI template metadata." `
            -TableName $TableName `
            -FieldName $columnName

        if ($codeExamples.Count -gt 0) {
            $examples = @($codeExamples | Select-Object -First 20)
            $exampleText = [string]::Join("; ", [string[]]$examples)
            if ($codeExamples.Count -gt $examples.Count) {
                $exampleText += "; ..."
            }
            $message += " Valid code examples: $exampleText"
        }

        $auditCount += Add-ConditionAudit `
            -Connection $Connection `
            -Transaction $Transaction `
            -TableName $TableName `
            -RuleName "Invalid CFI code" `
            -FieldName $columnName `
            -ObservedExpression $observedExpression `
            -Message $message `
            -RecordLabelExpression $recordLabel `
            -Condition $condition
    }

    return $auditCount
}

function Add-RequiredRegenSpeciesChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    if (-not (Test-ColumnExists -Columns $Columns -Name "SpeciesCode")) { return 0 }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "RegenMeasurements" -FieldName "SpeciesCode")) { return 0 }

    $condition = Get-RegenSpeciesRequiredCondition -Columns $Columns -SpeciesField "[SpeciesCode]"
    if ([string]::IsNullOrWhiteSpace($condition)) { return 0 }

    $plotField = Find-CandidateColumn -Columns $Columns -Patterns @("^plotmeaskey$", "^plotkey$", "^plotid$", "^plotnumber$", "^plot$", "plotid", "plotnumber", "stand", "unit")
    $treeField = Find-CandidateColumn -Columns $Columns -Patterns @("^regenmeaskey$", "^treemeaskey$", "^treekey$", "^treeid$", "^treenumber$", "^tree$", "tag", "stem")
    $recordLabel = Get-RecordLabelExpression -PlotField $plotField -TreeField $treeField

    $observedParts = New-Object System.Collections.Generic.List[string]
    [void]$observedParts.Add("'SpeciesCode is blank'")
    foreach ($fieldName in @("RegenMeasKey", "IDBH", "StemCount", "MinorPlot")) {
        if (-not (Test-ColumnExists -Columns $Columns -Name $fieldName)) { continue }
        [void]$observedParts.Add("'; $fieldName=' & ($(Quote-Name $fieldName) & '')")
    }
    $observedExpression = "(" + ([string]::Join(" & ", [string[]]$observedParts.ToArray())) + ")"

    $message = Add-FieldManualTipToMessage `
        -Message "RegenMeasurements.SpeciesCode is required when a regen row has RegenMeasKey, IDBH, StemCount, or MinorPlot data recorded." `
        -TableName "RegenMeasurements" `
        -FieldName "SpeciesCode"

    return Add-ConditionAudit `
        -Connection $Connection `
        -Transaction $null `
        -TableName "RegenMeasurements" `
        -RuleName "Required regen species" `
        -FieldName "SpeciesCode" `
        -ObservedExpression $observedExpression `
        -Message $message `
        -RecordLabelExpression $recordLabel `
        -Condition $condition
}

function Add-DuplicateAudit {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.OleDb.OleDbTransaction]$Transaction,
        [string]$TableName,
        [string]$PlotField,
        [string]$TreeField
    )

    $keyFields = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($PlotField)) {
        [void]$keyFields.Add($PlotField)
    }
    if (-not [string]::IsNullOrWhiteSpace($TreeField) -and -not $keyFields.Contains($TreeField)) {
        [void]$keyFields.Add($TreeField)
    }

    if ($keyFields.Count -eq 0) { return 0 }

    $quotedFields = @($keyFields | ForEach-Object { Quote-Name $_ })
    $notNullCondition = ($quotedFields | ForEach-Object { "$_ Is Not Null" }) -join " AND "
    $groupBy = $quotedFields -join ", "
    $label = ($keyFields | ForEach-Object { Field-Text-Expression $_ }) -join " & '/' & "
    $fieldLabel = [string]::Join("/", [string[]]$keyFields.ToArray())
    $joinCondition = ($quotedFields | ForEach-Object { "t.$_ = d.$_" }) -join " AND "
    $tempTableName = "InventoryCleanDup_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $tempTable = Quote-Name $tempTableName

    try {
        $selectIntoSql = @"
SELECT $groupBy
INTO $tempTable
FROM $(Quote-Name $TableName)
WHERE $notNullCondition
GROUP BY $groupBy
HAVING Count(*) > 1
"@
        [void](Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql $selectIntoSql)

        $auditSql = @"
INSERT INTO [InventoryCleanAudit]
    ([AuditTime], [TableName], [RuleName], [RecordLabel], [FieldName], [ObservedValue], [Message])
SELECT
    Now(),
    $(Sql-Text $TableName),
    'Duplicate selected key',
    $label,
    $(Sql-Text $fieldLabel),
    '',
    'More than one row uses the same selected key field or fields.'
FROM $tempTable
"@
        $rows = Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql $auditSql

        $markSql = @"
UPDATE $(Quote-Name $TableName) AS t
INNER JOIN $tempTable AS d
ON $joinCondition
SET t.[NeedsReview] = True
"@
        [void](Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql $markSql)

        return $rows
    }
    finally {
        try {
            [void](Invoke-NonQuery -Connection $Connection -Transaction $Transaction -Sql "DROP TABLE $tempTable")
        }
        catch {
        }
    }
}

function Backup-Database {
    param([string]$Path)

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [System.IO.Path]::GetExtension($Path)
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupPath = Join-Path $directory "$name.cleaned-backup-$timestamp$extension"

    Copy-Item -LiteralPath $Path -Destination $backupPath -ErrorAction Stop
    return $backupPath
}

function Test-TableAvailable {
    param(
        [string[]]$Tables,
        [string]$TableName
    )

    return $Tables -contains $TableName
}

function Get-CfiWorkbookTargetTables {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $existingTables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $targets = @(
        "Plots",
        "PlotMeasurements",
        "PlotCustomMeasurements",
        "Trees",
        "TreeMeasurements",
        "TreeCustomMeasurements",
        "RegenMeasurements",
        "RegenCustomMeasurements"
    )

    return @($targets | Where-Object { Test-TableAvailable -Tables $existingTables -TableName $_ })
}

function Add-PeriodNumber {
    param(
        [System.Collections.Generic.List[int]]$Periods,
        [object]$Value
    )

    if ($null -eq $Value -or [System.DBNull]::Value.Equals($Value)) { return }
    $period = [int]$Value
    if (-not $Periods.Contains($period)) {
        [void]$Periods.Add($period)
    }
}

function Get-PeriodNumbers {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $periods = New-Object System.Collections.Generic.List[int]

    if (Test-TableAvailable -Tables $tables -TableName "ProjectMeasurementPeriods") {
        $periodTable = Get-DataTable -Connection $Connection -Sql "SELECT [PeriodNumber] FROM [ProjectMeasurementPeriods] WHERE [PeriodNumber] Is Not Null ORDER BY [PeriodNumber]"
        foreach ($row in $periodTable.Rows) {
            Add-PeriodNumber -Periods $periods -Value $row["PeriodNumber"]
        }
    }

    foreach ($tableName in @("PlotMeasurements", "TreeMeasurements")) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $tableName)) { continue }
        $columns = @(Get-TableColumns -Connection $Connection -TableName $tableName)
        if (-not (Test-ColumnExists -Columns $columns -Name "PeriodNumber")) { continue }

        $periodTable = Get-DataTable -Connection $Connection -Sql "SELECT DISTINCT [PeriodNumber] FROM $(Quote-Name $tableName) WHERE [PeriodNumber] Is Not Null"
        foreach ($row in $periodTable.Rows) {
            Add-PeriodNumber -Periods $periods -Value $row["PeriodNumber"]
        }
    }

    return @($periods | Sort-Object)
}

function Get-ActiveWhereClause {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    if (Test-ColumnExists -Columns $columns -Name "IsDeleted") {
        return " WHERE ([IsDeleted] = False OR [IsDeleted] Is Null)"
    }

    return ""
}

function Get-CountValue {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$Sql
    )

    $value = Get-Scalar -Connection $Connection -Sql $Sql
    if ($null -eq $value -or [System.DBNull]::Value.Equals($value)) { return 0 }
    return [int]$value
}

function Get-RegenMeasKeyPeriodCondition {
    param(
        [int]$PeriodNumber,
        [string]$Alias = ""
    )

    $periodExpression = Get-RegenMeasKeyPeriodValueExpression -Alias $Alias
    $periodText = $PeriodNumber.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    return "($periodExpression = $periodText)"
}

function Get-RegenMeasKeyPeriodValueExpression {
    param([string]$Alias = "")

    $field = if ([string]::IsNullOrWhiteSpace($Alias)) { "[RegenMeasKey]" } else { "$Alias.[RegenMeasKey]" }
    return "Val(Mid(($field & ''), InStr(1, ($field & ''), '-') + 1, 2))"
}

function Get-RegenKeyCountVerificationSql {
    param([int]$PeriodNumber = 0)

    $regenPeriodExpression = Get-RegenMeasKeyPeriodValueExpression
    $customPeriodExpression = Get-RegenMeasKeyPeriodValueExpression
    $regenWhere = "Trim(([RegenMeasKey] & '')) <> ''"
    $customWhere = "Trim(([RegenMeasKey] & '')) <> ''"
    if ($PeriodNumber -gt 0) {
        $regenWhere += " AND " + (Get-RegenMeasKeyPeriodCondition -PeriodNumber $PeriodNumber)
        $customWhere += " AND " + (Get-RegenMeasKeyPeriodCondition -PeriodNumber $PeriodNumber)
    }

    return @"
SELECT 'RegenMeasurements' AS [SourceTable], $regenPeriodExpression AS [PeriodNumber], Count(*) AS [RecordCount]
FROM [RegenMeasurements]
WHERE $regenWhere
GROUP BY $regenPeriodExpression
UNION ALL
SELECT 'RegenCustomMeasurements' AS [SourceTable], $customPeriodExpression AS [PeriodNumber], Count(*) AS [RecordCount]
FROM [RegenCustomMeasurements]
WHERE $customWhere
GROUP BY $customPeriodExpression
ORDER BY [PeriodNumber], [SourceTable];
"@
}

function Test-CleaningPeriodScopeEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["PeriodScopeEnabled"]) {
                return [bool]$settings.PeriodScopeEnabled
            }
        }
        catch {
        }
        return $false
    }

    try {
        return ($null -ne $periodScopeCheck -and $periodScopeCheck.Checked)
    }
    catch {
        return $false
    }
}

function Test-SingleMeasurementProjectEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["PeriodScopeEnabled"] -and (-not [bool]$settings.PeriodScopeEnabled)) {
                return $false
            }
            if ($settings.PSObject.Properties["SingleMeasurementProject"]) {
                return [bool]$settings.SingleMeasurementProject
            }
        }
        catch {
        }
        return $false
    }

    try {
        if ($null -ne $periodScopeCheck -and (-not $periodScopeCheck.Checked)) { return $false }
        return ($null -ne $singleMeasurementProjectCheck -and $singleMeasurementProjectCheck.Checked)
    }
    catch {
        return $false
    }
}

function Get-PeriodScopeCurrentPeriodValue {
    if (Test-SingleMeasurementProjectEnabled) { return 1 }

    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["CurrentPeriodNumber"]) {
                return [int]$settings.CurrentPeriodNumber
            }
        }
        catch {
        }
        return 0
    }

    try {
        if ($null -ne $currentPeriodBox) { return [int]$currentPeriodBox.Value }
    }
    catch {
    }

    return 0
}

function Get-PeriodScopePastPeriodValue {
    if (Test-SingleMeasurementProjectEnabled) { return 0 }

    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["PastPeriodNumber"]) {
                return [int]$settings.PastPeriodNumber
            }
        }
        catch {
        }
        return 0
    }

    try {
        if ($null -ne $pastPeriodBox) { return [int]$pastPeriodBox.Value }
    }
    catch {
    }

    return 0
}

function Get-SelectedCleaningPeriods {
    if (-not (Test-CleaningPeriodScopeEnabled)) { return @() }

    $periods = New-Object System.Collections.Generic.List[int]
    $periodValues = New-Object System.Collections.Generic.List[int]
    [void]$periodValues.Add((Get-PeriodScopeCurrentPeriodValue))
    if (-not (Test-SingleMeasurementProjectEnabled)) {
        [void]$periodValues.Add((Get-PeriodScopePastPeriodValue))
    }

    foreach ($period in @($periodValues.ToArray())) {
        if ($period -le 0) { continue }
        if (-not $periods.Contains([int]$period)) {
            [void]$periods.Add([int]$period)
        }
    }

    return @($periods.ToArray() | Sort-Object)
}

function Get-SelectedCleaningPeriodsSqlList {
    $periods = @(Get-SelectedCleaningPeriods)
    if ($periods.Count -eq 0) { return "" }

    $periodText = @($periods | ForEach-Object { ([int]$_).ToString([System.Globalization.CultureInfo]::InvariantCulture) })
    return [string]::Join(", ", [string[]]$periodText)
}

function Get-CleaningPeriodScopeDescription {
    if (-not (Test-CleaningPeriodScopeEnabled)) {
        return "Cleaning period scope: all periods; TreeHistory checked across all periods."
    }

    $current = Get-PeriodScopeCurrentPeriodValue
    $past = Get-PeriodScopePastPeriodValue
    $currentText = if ($current -gt 0) { [string]$current } else { "not selected" }
    if (Test-SingleMeasurementProjectEnabled) {
        return "Cleaning period scope: single measurement project, measurement period $currentText; TreeHistory checked across all periods."
    }

    $pastText = if ($past -gt 0) { [string]$past } else { "not selected" }
    return "Cleaning period scope: current period $currentText, previous period $pastText; TreeHistory checked across all periods."
}

function Test-AllPeriodTreeHistoryRule {
    param(
        [string]$RuleName,
        [string]$FieldName
    )

    if ($FieldName -and $FieldName.Trim().Equals("TreeHistory", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    return ($RuleName -in @("TreeHistory transition check", "Lazarus tree check"))
}

function Get-SelectedPeriodAliasCondition {
    param(
        [string]$Alias,
        [string]$FieldName = "PeriodNumber"
    )

    $periodList = Get-SelectedCleaningPeriodsSqlList
    if ([string]::IsNullOrWhiteSpace($periodList)) { return "" }

    $field = if ([string]::IsNullOrWhiteSpace($Alias)) { Quote-Name $FieldName } else { "$Alias.$(Quote-Name $FieldName)" }
    return "$field In ($periodList)"
}

function Get-CurrentMeasurementPeriodAliasCondition {
    param(
        [string]$Alias,
        [string]$FieldName = "PeriodNumber"
    )

    if (-not (Test-CleaningPeriodScopeEnabled)) { return "" }

    $currentPeriod = Get-PeriodScopeCurrentPeriodValue
    if ($currentPeriod -le 0) { return "" }

    $field = if ([string]::IsNullOrWhiteSpace($Alias)) { Quote-Name $FieldName } else { "$Alias.$(Quote-Name $FieldName)" }
    $currentText = ([int]$currentPeriod).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    return "$field = $currentText"
}

function Get-SelectedPreviousCurrentPairCondition {
    param(
        [string]$EarlierAlias,
        [string]$LaterAlias
    )

    if (-not (Test-CleaningPeriodScopeEnabled)) { return "" }
    if (Test-SingleMeasurementProjectEnabled) { return "1 = 0" }

    $currentPeriod = Get-PeriodScopeCurrentPeriodValue
    $previousPeriod = Get-PeriodScopePastPeriodValue
    if ($currentPeriod -le 0 -or $previousPeriod -le 0) { return "1 = 0" }

    $currentText = ([int]$currentPeriod).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $previousText = ([int]$previousPeriod).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    return "$LaterAlias.[PeriodNumber] = $currentText AND $EarlierAlias.[PeriodNumber] = $previousText"
}

function Get-RegenMeasKeyPeriodScopeCondition {
    param([string]$Alias = "")

    $periods = @(Get-SelectedCleaningPeriods)
    if ($periods.Count -eq 0) { return "" }

    $conditions = @($periods | ForEach-Object { Get-RegenMeasKeyPeriodCondition -PeriodNumber ([int]$_) -Alias $Alias })
    return "(" + ([string]::Join(" OR ", [string[]]$conditions)) + ")"
}

function Get-PeriodNumbersForCleaningScope {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $selectedPeriods = @(Get-SelectedCleaningPeriods)
    if ((Test-CleaningPeriodScopeEnabled) -and $selectedPeriods.Count -gt 0) {
        return @($selectedPeriods)
    }

    return @(Get-PeriodNumbers -Connection $Connection)
}

function Get-PeriodScopeConditionForTable {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$RuleName = "",
        [string]$FieldName = ""
    )

    if (-not (Test-CleaningPeriodScopeEnabled)) { return "" }
    if (Test-AllPeriodTreeHistoryRule -RuleName $RuleName -FieldName $FieldName) { return "" }

    $periodList = Get-SelectedCleaningPeriodsSqlList
    if ([string]::IsNullOrWhiteSpace($periodList)) { return "" }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    switch ($TableName) {
        "ProjectMeasurementPeriods" {
            return "[PeriodNumber] In ($periodList)"
        }
        "PlotMeasurements" {
            $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
            if (Test-ColumnExists -Columns $columns -Name "PeriodNumber") {
                return "[PeriodNumber] In ($periodList)"
            }
        }
        "TreeMeasurements" {
            $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
            if (Test-ColumnExists -Columns $columns -Name "PeriodNumber") {
                return "[PeriodNumber] In ($periodList)"
            }
        }
        "PlotCustomMeasurements" {
            if (-not (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements")) { return "" }
            $customColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotCustomMeasurements")
            $measurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
            if ((Test-ColumnExists -Columns $customColumns -Name "PlotMeasKey") -and
                (Test-ColumnExists -Columns $measurementColumns -Name "PlotMeasKey") -and
                (Test-ColumnExists -Columns $measurementColumns -Name "PeriodNumber")) {
                return "EXISTS (SELECT 1 FROM [PlotMeasurements] AS periodScopePlotMeas WHERE periodScopePlotMeas.[PlotMeasKey] = [PlotCustomMeasurements].[PlotMeasKey] AND periodScopePlotMeas.[PeriodNumber] In ($periodList))"
            }
        }
        "TreeCustomMeasurements" {
            if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return "" }
            $customColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeCustomMeasurements")
            $measurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
            if ((Test-ColumnExists -Columns $customColumns -Name "TreeMeasKey") -and
                (Test-ColumnExists -Columns $measurementColumns -Name "TreeMeasKey") -and
                (Test-ColumnExists -Columns $measurementColumns -Name "PeriodNumber")) {
                return "EXISTS (SELECT 1 FROM [TreeMeasurements] AS periodScopeTreeMeas WHERE periodScopeTreeMeas.[TreeMeasKey] = [TreeCustomMeasurements].[TreeMeasKey] AND periodScopeTreeMeas.[PeriodNumber] In ($periodList))"
            }
        }
        "RegenMeasurements" {
            $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
            if (Test-ColumnExists -Columns $columns -Name "PeriodNumber") {
                return "[PeriodNumber] In ($periodList)"
            }
            if ((Test-ColumnExists -Columns $columns -Name "ProjectPeriodID") -and
                (Test-TableAvailable -Tables $tables -TableName "ProjectMeasurementPeriods")) {
                $periodColumns = @(Get-TableColumns -Connection $Connection -TableName "ProjectMeasurementPeriods")
                if ((Test-ColumnExists -Columns $periodColumns -Name "ProjectPeriodID") -and
                    (Test-ColumnExists -Columns $periodColumns -Name "PeriodNumber")) {
                    return "EXISTS (SELECT 1 FROM [ProjectMeasurementPeriods] AS periodScopeRegenPeriod WHERE periodScopeRegenPeriod.[ProjectPeriodID] = [RegenMeasurements].[ProjectPeriodID] AND periodScopeRegenPeriod.[PeriodNumber] In ($periodList))"
                }
            }
            if (Test-ColumnExists -Columns $columns -Name "RegenMeasKey") {
                return Get-RegenMeasKeyPeriodScopeCondition
            }
        }
        "RegenCustomMeasurements" {
            $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
            if ((Test-ColumnExists -Columns $columns -Name "RegenMeasKey") -and
                (Test-TableAvailable -Tables $tables -TableName "RegenMeasurements")) {
                $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements")
                if ((Test-ColumnExists -Columns $regenColumns -Name "RegenMeasKey") -and
                    (Test-ColumnExists -Columns $regenColumns -Name "ProjectPeriodID") -and
                    (Test-TableAvailable -Tables $tables -TableName "ProjectMeasurementPeriods")) {
                    $periodColumns = @(Get-TableColumns -Connection $Connection -TableName "ProjectMeasurementPeriods")
                    if ((Test-ColumnExists -Columns $periodColumns -Name "ProjectPeriodID") -and
                        (Test-ColumnExists -Columns $periodColumns -Name "PeriodNumber")) {
                        return "EXISTS (SELECT 1 FROM ([RegenMeasurements] AS periodScopeRegen INNER JOIN [ProjectMeasurementPeriods] AS periodScopeRegenPeriod ON periodScopeRegen.[ProjectPeriodID] = periodScopeRegenPeriod.[ProjectPeriodID]) WHERE periodScopeRegen.[RegenMeasKey] = [RegenCustomMeasurements].[RegenMeasKey] AND periodScopeRegenPeriod.[PeriodNumber] In ($periodList))"
                    }
                }
            }
            if (Test-ColumnExists -Columns $columns -Name "RegenMeasKey") {
                return Get-RegenMeasKeyPeriodScopeCondition
            }
        }
        "Trees" {
            if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return "" }
            $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
            $measurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
            if ((Test-ColumnExists -Columns $treeColumns -Name "TreeID") -and
                (Test-ColumnExists -Columns $measurementColumns -Name "TreeID") -and
                (Test-ColumnExists -Columns $measurementColumns -Name "PeriodNumber")) {
                return "EXISTS (SELECT 1 FROM [TreeMeasurements] AS periodScopeTreeMeas WHERE periodScopeTreeMeas.[TreeID] = [Trees].[TreeID] AND periodScopeTreeMeas.[PeriodNumber] In ($periodList))"
            }
        }
        "Plots" {
            if (-not (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements")) { return "" }
            $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
            $measurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
            if ((Test-ColumnExists -Columns $plotColumns -Name "PlotID") -and
                (Test-ColumnExists -Columns $measurementColumns -Name "PlotID") -and
                (Test-ColumnExists -Columns $measurementColumns -Name "PeriodNumber")) {
                return "EXISTS (SELECT 1 FROM [PlotMeasurements] AS periodScopePlotMeas WHERE periodScopePlotMeas.[PlotID] = [Plots].[PlotID] AND periodScopePlotMeas.[PeriodNumber] In ($periodList))"
            }
        }
    }

    return ""
}

function Get-CachedPeriodScopeConditionForTable {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$RuleName = "",
        [string]$FieldName = ""
    )

    if (-not (Test-CleaningPeriodScopeEnabled)) { return "" }
    if (Test-AllPeriodTreeHistoryRule -RuleName $RuleName -FieldName $FieldName) { return "" }

    $periodList = Get-SelectedCleaningPeriodsSqlList
    if ([string]::IsNullOrWhiteSpace($periodList)) { return "" }

    if ($null -eq $script:PeriodScopeConditionCache) {
        $script:PeriodScopeConditionCache = @{}
    }

    $sourceKey = ""
    try { $sourceKey = [string]$Connection.DataSource } catch { $sourceKey = "" }
    $cacheKey = "$sourceKey|$TableName|$periodList"
    if ($script:PeriodScopeConditionCache.ContainsKey($cacheKey)) {
        return [string]$script:PeriodScopeConditionCache[$cacheKey]
    }

    $condition = Get-PeriodScopeConditionForTable -Connection $Connection -TableName $TableName -RuleName "" -FieldName ""
    $script:PeriodScopeConditionCache[$cacheKey] = [string]$condition
    return $condition
}

function Add-PeriodScopeToCondition {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$RuleName,
        [string]$FieldName,
        [string]$Condition
    )

    $scopeCondition = Get-CachedPeriodScopeConditionForTable -Connection $Connection -TableName $TableName -RuleName $RuleName -FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($scopeCondition)) { return $Condition }
    if ([string]::IsNullOrWhiteSpace($Condition)) { return $scopeCondition }

    return "(($Condition) AND ($scopeCondition))"
}

function Add-SelectedPeriodAliasConditionToWhereClause {
    param(
        [string]$WhereClause,
        [string]$Alias
    )

    $periodCondition = Get-SelectedPeriodAliasCondition -Alias $Alias
    if ([string]::IsNullOrWhiteSpace($periodCondition)) { return $WhereClause }
    if ([string]::IsNullOrWhiteSpace($WhereClause)) { return $periodCondition }

    return "(($WhereClause) AND ($periodCondition))"
}

function New-ScopedSelectVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$RuleName,
        [string]$FieldName,
        [string]$Condition,
        [string]$SelectList = "*"
    )

    $where = Add-PeriodScopeToCondition -Connection $Connection -TableName $TableName -RuleName $RuleName -FieldName $FieldName -Condition $Condition
    return "SELECT $SelectList FROM $(Quote-Name $TableName) WHERE $where;"
}

function Get-AliasSelectList {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$Alias
    )

    try {
        $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($column in $columns) {
            $name = [string]$column.Name
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            if ($name.Equals("NeedsReview", [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            [void]$parts.Add("$Alias.$(Quote-Name $name)")
        }
        if ($parts.Count -gt 0) {
            return [string]::Join(", ", [string[]]$parts.ToArray())
        }
    }
    catch {
    }

    return "*"
}

function Set-PeriodScopeDefaultsFromConnection {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    try {
        if ($null -eq $currentPeriodBox -or $null -eq $pastPeriodBox) { return }
        $periods = @(Get-PeriodNumbers -Connection $Connection)
        if ($periods.Count -eq 0) {
            $currentPeriodBox.Value = 0
            $pastPeriodBox.Value = 0
            if ($null -ne $singleMeasurementProjectCheck) { $singleMeasurementProjectCheck.Checked = $false }
            return
        }

        $current = [int]$periods[$periods.Count - 1]
        $past = 0
        $singleMeasurementProject = ($periods.Count -eq 1)
        if ($periods.Count -gt 1) {
            $past = [int]$periods[$periods.Count - 2]
        }
        else {
            $current = 1
        }

        $maxPeriod = [Math]::Max(999, [Math]::Max($current, $past))
        $currentPeriodBox.Maximum = $maxPeriod
        $pastPeriodBox.Maximum = $maxPeriod
        $currentPeriodBox.Value = $current
        $pastPeriodBox.Value = $past
        if ($null -ne $singleMeasurementProjectCheck) {
            $singleMeasurementProjectCheck.Checked = $singleMeasurementProject
        }
    }
    catch {
        Add-Log "Could not auto-fill current/past period values: $($_.Exception.Message)"
    }
}

function Add-CountMismatchAudit {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$RecordLabel,
        [int]$Expected,
        [int]$Observed,
        [string]$Message
    )

    if ($Expected -eq $Observed) { return 0 }

    Add-AuditRow `
        -Connection $Connection `
        -TableName $TableName `
        -RuleName "Workbook count check" `
        -RecordLabel $RecordLabel `
        -FieldName "Record count" `
        -ObservedValue "$Observed observed; $Expected expected" `
        -Message $Message

    return 1
}

function Add-WorkbookCountChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $auditCount = 0
    $periods = @(Get-PeriodNumbersForCleaningScope -Connection $Connection)
    if ($periods.Count -eq 0) { return 0 }

    if ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
        (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements")) {
        $plotCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [Plots]$(Get-ActiveWhereClause -Connection $Connection -TableName 'Plots')"
        foreach ($period in $periods) {
            $plotMeasCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [PlotMeasurements] WHERE [PeriodNumber] = $period"
            $auditCount += Add-CountMismatchAudit `
                -Connection $Connection `
                -TableName "PlotMeasurements" `
                -RecordLabel "Period $period" `
                -Expected $plotCount `
                -Observed $plotMeasCount `
                -Message "The spreadsheet says PlotMeasurements should match the plot record count for each measurement period."
        }
    }

    if ((Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") -and
        (Test-TableAvailable -Tables $tables -TableName "PlotCustomMeasurements")) {
        foreach ($period in $periods) {
            $plotMeasCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [PlotMeasurements] WHERE [PeriodNumber] = $period"
            $plotCustomCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [PlotCustomMeasurements] AS c INNER JOIN [PlotMeasurements] AS m ON c.[PlotMeasKey] = m.[PlotMeasKey] WHERE m.[PeriodNumber] = $period"
            $auditCount += Add-CountMismatchAudit `
                -Connection $Connection `
                -TableName "PlotCustomMeasurements" `
                -RecordLabel "Period $period" `
                -Expected $plotMeasCount `
                -Observed $plotCustomCount `
                -Message "The spreadsheet says PlotCustomMeasurements should match PlotMeasurements for each measurement period."
        }
    }

    if ((Test-TableAvailable -Tables $tables -TableName "Trees") -and
        (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) {
        $treeCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [Trees]$(Get-ActiveWhereClause -Connection $Connection -TableName 'Trees')"
        foreach ($period in $periods) {
            $treeMeasCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [TreeMeasurements] WHERE [PeriodNumber] = $period"
            $auditCount += Add-CountMismatchAudit `
                -Connection $Connection `
                -TableName "TreeMeasurements" `
                -RecordLabel "Period $period" `
                -Expected $treeCount `
                -Observed $treeMeasCount `
                -Message "The spreadsheet says TreeMeasurements should usually match the tree record count by period; review dropped or unmeasured plots when they differ."
        }
    }

    if ((Test-TableAvailable -Tables $tables -TableName "TreeMeasurements") -and
        (Test-TableAvailable -Tables $tables -TableName "TreeCustomMeasurements")) {
        foreach ($period in $periods) {
            $treeMeasCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [TreeMeasurements] WHERE [PeriodNumber] = $period"
            $treeCustomCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [TreeCustomMeasurements] AS c INNER JOIN [TreeMeasurements] AS m ON c.[TreeMeasKey] = m.[TreeMeasKey] WHERE m.[PeriodNumber] = $period"
            $auditCount += Add-CountMismatchAudit `
                -Connection $Connection `
                -TableName "TreeCustomMeasurements" `
                -RecordLabel "Period $period" `
                -Expected $treeMeasCount `
                -Observed $treeCustomCount `
                -Message "The spreadsheet says TreeCustomMeasurements should match TreeMeasurements for each measurement period."
        }
    }

    if ((Test-TableAvailable -Tables $tables -TableName "RegenMeasurements") -and
        (Test-TableAvailable -Tables $tables -TableName "RegenCustomMeasurements")) {
        $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements")
        $regenCustomColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenCustomMeasurements")
        if ((Test-ColumnExists -Columns $regenColumns -Name "RegenMeasKey") -and
            (Test-ColumnExists -Columns $regenCustomColumns -Name "RegenMeasKey")) {
            foreach ($period in $periods) {
                $regenCondition = Get-RegenMeasKeyPeriodCondition -PeriodNumber ([int]$period)
                $regenCustomCondition = Get-RegenMeasKeyPeriodCondition -PeriodNumber ([int]$period)
                $regenMeasCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [RegenMeasurements] WHERE $regenCondition"
                $regenCustomCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [RegenCustomMeasurements] WHERE $regenCustomCondition"
                $auditCount += Add-CountMismatchAudit `
                    -Connection $Connection `
                    -TableName "RegenCustomMeasurements" `
                    -RecordLabel "Period $period" `
                    -Expected $regenMeasCount `
                    -Observed $regenCustomCount `
                    -Message "The spreadsheet says RegenCustomMeasurements should match RegenMeasurements for each measurement period where regen records exist. Regen period counts are filtered from the RegenMeasKey in each table, for example P005337-01-3-122-040 is period 1."
            }
        }
    }

    return $auditCount
}

function Add-ProjectBuildAudit {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$RecordLabel,
        [string]$FieldName,
        [string]$ObservedValue,
        [string]$Message
    )

    Add-AuditRow `
        -Connection $Connection `
        -TableName "Project setup" `
        -RuleName "Project user guide check" `
        -RecordLabel $RecordLabel `
        -FieldName $FieldName `
        -ObservedValue $ObservedValue `
        -Message $Message

    return 1
}

function Get-DatabaseObjectDefinitions {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $definitions = @{}
    foreach ($schemaGuid in @([System.Data.OleDb.OleDbSchemaGuid]::Views, [System.Data.OleDb.OleDbSchemaGuid]::Procedures)) {
        try {
            $schema = $Connection.GetOleDbSchemaTable($schemaGuid, $null)
        }
        catch {
            continue
        }

        if ($null -eq $schema) { continue }

        foreach ($row in $schema.Rows) {
            $name = ""
            $definition = ""
            if ($schema.Columns.Contains("TABLE_NAME")) {
                $name = [string]$row["TABLE_NAME"]
            }
            if ([string]::IsNullOrWhiteSpace($name) -and $schema.Columns.Contains("PROCEDURE_NAME")) {
                $name = [string]$row["PROCEDURE_NAME"]
            }
            if ($schema.Columns.Contains("VIEW_DEFINITION")) {
                $definition = [string]$row["VIEW_DEFINITION"]
            }
            if ([string]::IsNullOrWhiteSpace($definition) -and $schema.Columns.Contains("PROCEDURE_DEFINITION")) {
                $definition = [string]$row["PROCEDURE_DEFINITION"]
            }

            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $definitions[$name] = $definition
            }
        }
    }

    return $definitions
}

function Add-ProjectSetupChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $auditCount = 0
    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)

    $requiredTables = @(
        "Projects",
        "ProjectMeasurementPeriods",
        "Plots",
        "PlotMeasurements",
        "PlotCustomMeasurements",
        "Trees",
        "TreeMeasurements",
        "TreeCustomMeasurements",
        "RegenMeasurements",
        "RegenCustomMeasurements",
        "AppTables",
        "AppColumns",
        "AppColumnCodes",
        "FieldDisplayRules",
        "FieldDisplayRuleChecks",
        "InventoryAssignments",
        "InventoryAssignmentPlots",
        "InventoryAssignmentTrees",
        "ReportHeaders",
        "ReportColumns",
        "ValidationRules",
        "AnalysisProjectGenerators",
        "AnalysisProjectParameters",
        "AnalysisRuns"
    )

    foreach ($tableName in $requiredTables) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $tableName)) {
            $auditCount += Add-ProjectBuildAudit `
                -Connection $Connection `
                -RecordLabel $tableName `
                -FieldName "Required table" `
                -ObservedValue "Missing" `
                -Message "The project build guide expects this table to exist in a completed CFI project database."
        }
    }

    if (Test-TableAvailable -Tables $tables -TableName "Projects") {
        $projectCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [Projects]"
        if ($projectCount -ne 1) {
            $auditCount += Add-ProjectBuildAudit `
                -Connection $Connection `
                -RecordLabel "Projects" `
                -FieldName "ProjectID" `
                -ObservedValue "$projectCount project records" `
                -Message "The project build guide describes a single unique ProjectID for the project; review the Projects table."
        }

        $placeholderCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [Projects] WHERE InStr(1, ([Description] & ''), 'Delete Me', 1) > 0"
        if ($placeholderCount -gt 0) {
            $auditCount += Add-ProjectBuildAudit `
                -Connection $Connection `
                -RecordLabel "Projects" `
                -FieldName "Description" `
                -ObservedValue "$placeholderCount placeholder records" `
                -Message "The project build guide says sample project text should be replaced during setup."
        }
    }

    if (Test-TableAvailable -Tables $tables -TableName "ProjectMeasurementPeriods") {
        $badPeriodCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [ProjectMeasurementPeriods] WHERE [ProjectID] Is Null OR [PeriodNumber] Is Null OR [PeriodDescription] Is Null OR InStr(1, ([PeriodDescription] & ''), 'Delete Me', 1) > 0"
        if ($badPeriodCount -gt 0) {
            $auditCount += Add-ProjectBuildAudit `
                -Connection $Connection `
                -RecordLabel "ProjectMeasurementPeriods" `
                -FieldName "Period setup" `
                -ObservedValue "$badPeriodCount records need review" `
                -Message "The project build guide says each measurement period should have the correct ProjectID, period number, and project-specific period description."
        }
    }

    foreach ($tableName in @("ProjectMeasurementPeriods", "Plots", "AnalysisProjectGenerators", "AnalysisProjectParameters", "AnalysisRuns", "ValidationRules", "FieldDisplayRules")) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $tableName)) { continue }
        $columns = @(Get-TableColumns -Connection $Connection -TableName $tableName)
        if (-not (Test-ColumnExists -Columns $columns -Name "ProjectID")) { continue }

        $badProjectIdCount = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM $(Quote-Name $tableName) AS t LEFT JOIN [Projects] AS p ON t.[ProjectID] = p.[ProjectID] WHERE t.[ProjectID] Is Null OR p.[ProjectID] Is Null"
        if ($badProjectIdCount -gt 0) {
            $auditCount += Add-ProjectBuildAudit `
                -Connection $Connection `
                -RecordLabel $tableName `
                -FieldName "ProjectID" `
                -ObservedValue "$badProjectIdCount unmatched records" `
                -Message "The project build guide says the new ProjectID must be carried into this table so the database links correctly."
        }
    }

    foreach ($tableName in @("PlotCustomMeasurements", "TreeCustomMeasurements", "RegenCustomMeasurements")) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $tableName)) { continue }
        $columns = @(Get-TableColumns -Connection $Connection -TableName $tableName)
        if (Test-ColumnExists -Columns $columns -Name "FieldNameSampleRemove") {
            $auditCount += Add-ProjectBuildAudit `
                -Connection $Connection `
                -RecordLabel $tableName `
                -FieldName "FieldNameSampleRemove" `
                -ObservedValue "Placeholder field still present" `
                -Message "The project build guide says to delete the sample custom-measurement field and replace it with the fields collected or archived for this project."
        }
    }

    if ((Test-TableAvailable -Tables $tables -TableName "AppTables") -and
        (Test-TableAvailable -Tables $tables -TableName "AppColumns")) {
        $appColumns = Get-DataTable -Connection $Connection -Sql "SELECT t.[TableName], c.[ColumnName], c.[Active] FROM [AppTables] AS t INNER JOIN [AppColumns] AS c ON t.[ID] = c.[TableID] WHERE c.[Active] = True"
        foreach ($row in $appColumns.Rows) {
            $tableName = [string]$row["TableName"]
            $columnName = [string]$row["ColumnName"]
            if (-not (Test-TableAvailable -Tables $tables -TableName $tableName)) { continue }

            $columns = @(Get-TableColumns -Connection $Connection -TableName $tableName)
            if (-not (Test-ColumnExists -Columns $columns -Name $columnName)) {
                $auditCount += Add-ProjectBuildAudit `
                    -Connection $Connection `
                    -RecordLabel $tableName `
                    -FieldName $columnName `
                    -ObservedValue "Active AppColumns entry without table field" `
                    -Message "The project build guide says AppColumns ColumnName must match the actual table field name."
            }
        }
    }

    $definitions = Get-DatabaseObjectDefinitions -Connection $Connection
    $requiredQueries = @(
        "_prjCFIDEER_GetPlotMeasurementsForPeriod",
        "_prjCFIDEER_GetTreeMeasurementsForPeriodByKey",
        "_prjCFIDEER_GetRegenMeasurements",
        "_prjCFIDEER_GetProcessPlotList",
        "_prjCFIDEER_GetProcessTreeList",
        "_prjCFIDEER_GetProcessRegenList",
        "_prjCFIDEER_UpdatePlotMeasurement",
        "_prjCFIDEER_UpdateTreeMeasurement",
        "_prjCFIDEER_UpdateRegenMeasurement"
    )

    foreach ($queryName in $requiredQueries) {
        if (-not $definitions.ContainsKey($queryName)) {
            $auditCount += Add-ProjectBuildAudit `
                -Connection $Connection `
                -RecordLabel $queryName `
                -FieldName "Required query" `
                -ObservedValue "Missing" `
                -Message "The project build guide identifies this get/update query as part of the CFI desktop application setup."
            continue
        }

        $definition = [string]$definitions[$queryName]
        if ($definition -match "FieldNameSampleRemove|Expr1") {
            $auditCount += Add-ProjectBuildAudit `
                -Connection $Connection `
                -RecordLabel $queryName `
                -FieldName "Query placeholder" `
                -ObservedValue "Placeholder still present" `
                -Message "The project build guide says the Expr1/sample custom-measurement placeholder should be replaced or deleted when the get/update query is set up."
        }
    }

    return $auditCount
}

function Get-ControlDecimalValue {
    param(
        [object]$Control,
        [decimal]$DefaultValue
    )

    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($null -eq $Control) {
                switch ([decimal]$DefaultValue) {
                    500 { return [decimal]$settings.DbhMax }
                    3500 { return [decimal]$settings.DbhMax }
                    250 { return [decimal]$settings.HeightMax }
                    150 { return [decimal]$settings.HeightMax }
                    999 { return [decimal]$settings.StemCountMax }
                    100 { return [decimal]$settings.DbhGrowthMax }
                    15 { return [decimal]$settings.DbhGrowthMax }
                    60 { return [decimal]$settings.HeightGrowthMax }
                    20 { return [decimal]$settings.HeightGrowthMax }
                }
            }

            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $dbhMax)) { return [decimal]$settings.DbhMax }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $heightMax)) { return [decimal]$settings.HeightMax }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $stemCountMax)) { return [decimal]$settings.StemCountMax }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $dbhGrowthMax)) { return [decimal]$settings.DbhGrowthMax }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $heightGrowthMax)) { return [decimal]$settings.HeightGrowthMax }
        }
        catch {
        }
    }

    if ($null -ne $Control -and $null -ne $Control.Value) {
        return [decimal]$Control.Value
    }

    return $DefaultValue
}

function Get-StemCountMaximumValue {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings -and $settings.PSObject.Properties["StemCountMax"]) {
        try { return [decimal]$settings.StemCountMax } catch { }
    }

    try {
        if ($null -ne $stemCountMax -and $null -ne $stemCountMax.Value) {
            return [decimal]$stemCountMax.Value
        }
    }
    catch {
    }

    return [decimal]100
}

function Get-ControlTextValue {
    param(
        [object]$Control,
        [string]$DefaultValue,
        [string]$SettingName = ""
    )

    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($SettingName) -and $settings.PSObject.Properties[$SettingName]) {
                return [string]$settings.PSObject.Properties[$SettingName].Value
            }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $woodlandSpeciesCodesBox)) {
                return [string]$settings.WoodlandSpeciesCodesText
            }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $heightRequiredMinorPlotsBox)) {
                if ($settings.PSObject.Properties["HeightRequiredMinorPlotsText"]) {
                    return [string]$settings.HeightRequiredMinorPlotsText
                }
            }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $regenTimberSeedlingMinorPlotsBox)) {
                if ($settings.PSObject.Properties["RegenTimberSeedlingMinorPlotsText"]) {
                    return [string]$settings.RegenTimberSeedlingMinorPlotsText
                }
            }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $regenTimberSapling20MinorPlotsBox)) {
                if ($settings.PSObject.Properties["RegenTimberSapling20MinorPlotsText"]) {
                    return [string]$settings.RegenTimberSapling20MinorPlotsText
                }
                if ($settings.PSObject.Properties["RegenTimberSaplingMinorPlotsText"]) {
                    return [string]$settings.RegenTimberSaplingMinorPlotsText
                }
            }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $regenTimberSapling40MinorPlotsBox)) {
                if ($settings.PSObject.Properties["RegenTimberSapling40MinorPlotsText"]) {
                    return [string]$settings.RegenTimberSapling40MinorPlotsText
                }
                if ($settings.PSObject.Properties["RegenTimberSaplingMinorPlotsText"]) {
                    return [string]$settings.RegenTimberSaplingMinorPlotsText
                }
            }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $regenWoodlandSeedlingMinorPlotsBox)) {
                if ($settings.PSObject.Properties["RegenWoodlandSeedlingMinorPlotsText"]) {
                    return [string]$settings.RegenWoodlandSeedlingMinorPlotsText
                }
            }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $regenWoodlandSapling20MinorPlotsBox)) {
                if ($settings.PSObject.Properties["RegenWoodlandSapling20MinorPlotsText"]) {
                    return [string]$settings.RegenWoodlandSapling20MinorPlotsText
                }
                if ($settings.PSObject.Properties["RegenWoodlandSaplingMinorPlotsText"]) {
                    return [string]$settings.RegenWoodlandSaplingMinorPlotsText
                }
            }
            if ($null -ne $Control -and [object]::ReferenceEquals($Control, $regenWoodlandSapling40MinorPlotsBox)) {
                if ($settings.PSObject.Properties["RegenWoodlandSapling40MinorPlotsText"]) {
                    return [string]$settings.RegenWoodlandSapling40MinorPlotsText
                }
                if ($settings.PSObject.Properties["RegenWoodlandSaplingMinorPlotsText"]) {
                    return [string]$settings.RegenWoodlandSaplingMinorPlotsText
                }
            }
        }
        catch {
        }
    }

    try {
        if ($null -ne $Control -and $null -ne $Control.Text) {
            return [string]$Control.Text
        }
    }
    catch {
    }

    return $DefaultValue
}

function Test-RangeChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        return [bool]$settings.RangeChecksEnabled
    }

    try {
        return ($null -ne $rangeCheck -and $rangeCheck.Checked)
    }
    catch {
        return $true
    }
}

function Test-TreeStemCountChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["TreeStemCountChecksEnabled"]) {
                return [bool]$settings.TreeStemCountChecksEnabled
            }
        }
        catch {
        }
        return $false
    }

    try {
        return ($null -ne $treeStemCountCheck -and $treeStemCountCheck.Checked)
    }
    catch {
        return $false
    }
}

function Test-NewMortalityIdbhChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["NewMortalityIdbhChecksEnabled"]) {
                return [bool]$settings.NewMortalityIdbhChecksEnabled
            }
        }
        catch {
        }
        return $false
    }

    try {
        return ($null -ne $newMortalityIdbhCheck -and $newMortalityIdbhCheck.Checked)
    }
    catch {
        return $false
    }
}

function Test-OldMortalityIdbhChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["OldMortalityIdbhChecksEnabled"]) {
                return [bool]$settings.OldMortalityIdbhChecksEnabled
            }
        }
        catch {
        }
        return $false
    }

    try {
        return ($null -ne $oldMortalityIdbhCheck -and $oldMortalityIdbhCheck.Checked)
    }
    catch {
        return $false
    }
}

function Test-NewMortalityHeightChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["NewMortalityHeightChecksEnabled"]) {
                return [bool]$settings.NewMortalityHeightChecksEnabled
            }
        }
        catch {
        }
        return $false
    }

    try {
        return ($null -ne $newMortalityHeightCheck -and $newMortalityHeightCheck.Checked)
    }
    catch {
        return $false
    }
}

function Test-OldMortalityHeightChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["OldMortalityHeightChecksEnabled"]) {
                return [bool]$settings.OldMortalityHeightChecksEnabled
            }
        }
        catch {
        }
        return $false
    }

    try {
        return ($null -ne $oldMortalityHeightCheck -and $oldMortalityHeightCheck.Checked)
    }
    catch {
        return $false
    }
}

function Test-Problem127HeightChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["Problem127HeightChecksEnabled"]) {
                return [bool]$settings.Problem127HeightChecksEnabled
            }
        }
        catch {
        }
        return $true
    }

    try {
        return ($null -ne $problem127HeightCheck -and $problem127HeightCheck.Checked)
    }
    catch {
        return $true
    }
}

function Test-Problem128HeightChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["Problem128HeightChecksEnabled"]) {
                return [bool]$settings.Problem128HeightChecksEnabled
            }
        }
        catch {
        }
        return $true
    }

    try {
        return ($null -ne $problem128HeightCheck -and $problem128HeightCheck.Checked)
    }
    catch {
        return $true
    }
}

function Test-Problem123HeightChecksEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["Problem123HeightChecksEnabled"]) {
                return [bool]$settings.Problem123HeightChecksEnabled
            }
        }
        catch {
        }
        return $true
    }

    try {
        return ($null -ne $problem123HeightCheck -and $problem123HeightCheck.Checked)
    }
    catch {
        return $true
    }
}

function Get-ProblemHeightNoHeightCodes {
    $codes = New-Object System.Collections.Generic.List[string]
    if (Test-Problem127HeightChecksEnabled) {
        Add-UniqueCode -Codes $codes -Code "127"
        Add-UniqueCode -Codes $codes -Code "74"
    }
    if (Test-Problem128HeightChecksEnabled) {
        Add-UniqueCode -Codes $codes -Code "128"
        Add-UniqueCode -Codes $codes -Code "75"
    }
    if (Test-Problem123HeightChecksEnabled) {
        Add-UniqueCode -Codes $codes -Code "123"
        Add-UniqueCode -Codes $codes -Code "72"
    }
    return [string[]]$codes.ToArray()
}

function Get-ProblemCodeDisplayText {
    param([string]$Code)

    $codeText = ([string]$Code).Trim()
    switch ($codeText) {
        "121" { return "121 (negative diameter growth)" }
        "127" { return "127 (broken/missing top)" }
        "74" { return "74 (legacy broken top)" }
        "128" { return "128 (dead top)" }
        "75" { return "75 (legacy dead top)" }
        "123" { return "123 (lean > 15 degrees)" }
        "72" { return "72 (legacy lean > 15 degrees)" }
        default { return $codeText }
    }
}

function Get-TimberProblemHeightCodeDisplayText {
    param([string]$Code)

    $codeText = ([string]$Code).Trim()
    switch ($codeText) {
        "127" { return "timber 127 (broken/missing top)" }
        "74" { return "timber 74 (legacy broken top)" }
        "128" { return "timber 128 (dead top)" }
        "75" { return "timber 75 (legacy dead top)" }
        "123" { return "timber 123 (lean > 15 degrees)" }
        "72" { return "timber 72 (legacy lean > 15 degrees)" }
        default { return $codeText }
    }
}

function Get-ProblemCodeLabelExpression {
    param([string]$ValueExpression)

    $textExpression = "Trim(($ValueExpression & ''))"
    $expression = "($ValueExpression & '')"
    foreach ($code in @("72", "74", "75", "123", "128", "127", "121")) {
        $expression = "IIf($textExpression = $(Sql-Text $code), $(Sql-Text (Get-ProblemCodeDisplayText -Code $code)), $expression)"
    }
    return $expression
}

function Get-ProblemHeightNoHeightCodesText {
    $codes = @(Get-ProblemHeightNoHeightCodes)
    if ($codes.Count -eq 0) { return "none selected" }
    return [string]::Join(", ", [string[]]@($codes | ForEach-Object { Get-TimberProblemHeightCodeDisplayText -Code $_ }))
}

function Get-TotalHeightProtocolMode {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["TotalHeightProtocolMode"]) {
                $mode = [string]$settings.TotalHeightProtocolMode
                if (-not [string]::IsNullOrWhiteSpace($mode)) { return $mode }
            }
        }
        catch {
        }
        return "Normal"
    }

    try {
        if ($null -ne $heightProtocolCombo -and $null -ne $heightProtocolCombo.SelectedItem) {
            $selected = [string]$heightProtocolCombo.SelectedItem
            switch -Wildcard ($selected) {
                "100%*" { return "FullLive" }
                "Subsample*" { return "Subsample" }
                default { return "Normal" }
            }
        }
    }
    catch {
    }

    return "Normal"
}

function Test-TotalHeightFullLiveProtocolEnabled {
    return (Get-TotalHeightProtocolMode) -eq "FullLive"
}

function Test-TotalHeightSubsampleProtocolEnabled {
    return (Get-TotalHeightProtocolMode) -eq "Subsample"
}

function Get-TotalHeightProtocolDisplayText {
    switch (Get-TotalHeightProtocolMode) {
        "FullLive" { return "100% live-tree heights for TreeHistory 0, 5, and 10" }
        "Subsample" { return "Subsample by plot/species/2-inch IDBH class; rare species still require 100% live-tree heights" }
        default { return "Normal required-height rules" }
    }
}

function Get-HeightProtocolNumericSetting {
    param(
        [System.Windows.Forms.NumericUpDown]$Control,
        [string]$SettingName,
        [decimal]$DefaultValue
    )

    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties[$SettingName]) {
                return [decimal]$settings.$SettingName
            }
        }
        catch {
        }
        return $DefaultValue
    }

    return (Get-ControlDecimalValue -Control $Control -DefaultValue $DefaultValue)
}

function Get-HeightSubsampleMinimumCount {
    return [int](Get-HeightProtocolNumericSetting -Control $heightSubsampleMinimumBox -SettingName "HeightSubsampleMinimumCount" -DefaultValue 2)
}

function Get-HeightSubsampleMinimumIdbh {
    return [decimal](Get-HeightProtocolNumericSetting -Control $heightSubsampleMinIdbhBox -SettingName "HeightSubsampleMinimumIdbh" -DefaultValue 50)
}

function Get-HeightSubsampleAllAtOrAboveIdbh {
    return [decimal](Get-HeightProtocolNumericSetting -Control $heightSubsampleAllAtOrAboveBox -SettingName "HeightSubsampleAllAtOrAboveIdbh" -DefaultValue 170)
}

function Test-HeightSubsampleAllAtOrAboveEnabled {
    $settings = $script:GuideRunSettings
    if ($null -ne $settings) {
        try {
            if ($settings.PSObject.Properties["HeightSubsampleAllAtOrAboveEnabled"]) {
                return [bool]$settings.HeightSubsampleAllAtOrAboveEnabled
            }
        }
        catch {
        }
        return $false
    }

    try {
        return ($null -ne $heightSubsampleAllAtOrAboveCheck -and $heightSubsampleAllAtOrAboveCheck.Checked)
    }
    catch {
        return $false
    }
}

function Get-HeightRareSpeciesCodeValues {
    $rawText = Get-ControlTextValue -Control $heightRareSpeciesCodesBox -DefaultValue "" -SettingName "HeightRareSpeciesCodesText"
    return @(Convert-DelimitedCodeTextToValues -Text $rawText)
}

function Get-HeightRareSpeciesCodesSqlList {
    $codes = @(Get-HeightRareSpeciesCodeValues)
    return (Convert-CodeValuesToSqlList -Values $codes)
}

function Get-HeightRareSpeciesDisplayText {
    $codes = @(Get-HeightRareSpeciesCodeValues | Sort-Object -Unique)
    if ($codes.Count -eq 0) { return "none entered" }
    return [string]::Join(", ", [string[]]$codes)
}

function Get-RequiredTreeIdbhHistoryCodes {
    $codes = New-Object System.Collections.Generic.List[string]
    foreach ($code in @("0", "10")) {
        Add-UniqueCode -Codes $codes -Code $code
    }

    if (Test-NewMortalityIdbhChecksEnabled) {
        foreach ($code in @("2", "3")) {
            Add-UniqueCode -Codes $codes -Code $code
        }
    }

    if (Test-OldMortalityIdbhChecksEnabled) {
        Add-UniqueCode -Codes $codes -Code "7"
    }

    return [string[]]$codes.ToArray()
}

function Get-TreeIdbhRequiredHistoryText {
    $codes = @(Get-RequiredTreeIdbhHistoryCodes)
    if ($codes.Count -eq 0) { return "no TreeHistory statuses" }
    return "TreeHistory " + ([string]::Join(", ", [string[]]$codes))
}

function Get-TreeIdbhMissingRequiredCondition {
    param(
        [string]$TreeHistoryField = "[TreeHistory]",
        [string]$IdbhField = "[IDBH]"
    )

    $historyValue = "Trim(($TreeHistoryField & ''))"
    $historyCodes = @(Get-RequiredTreeIdbhHistoryCodes)
    if ($historyCodes.Count -eq 0) { return "(1 = 0)" }

    $historyRequiredParts = New-Object System.Collections.Generic.List[string]

    if ($historyCodes.Count -gt 0) {
        $historyList = ([string[]]@($historyCodes | ForEach-Object { Sql-Text $_ })) -join ", "
        [void]$historyRequiredParts.Add("$historyValue In ($historyList)")
    }

    $historyRequired = [string]::Join(" OR ", [string[]]$historyRequiredParts.ToArray())
    return "(($historyRequired) AND $IdbhField Is Null)"
}

function Get-TreeIdbhShouldBeBlankCondition {
    param(
        [string]$TreeHistoryField = "[TreeHistory]",
        [string]$IdbhField = "[IDBH]"
    )

    $historyValue = "Trim(($TreeHistoryField & ''))"
    $blankHistoryList = Convert-CodeValuesToSqlList -Values @("1", "4", "8")
    return "($historyValue In ($blankHistoryList) AND $IdbhField Is Not Null AND Len(Trim(($IdbhField & ''))) > 0)"
}

function Get-TreeSpeciesRequiredCondition {
    param(
        [string]$TreeHistoryField = "tm.[TreeHistory]",
        [string]$SpeciesField = "treeMeta.[SpeciesCode]"
    )

    $speciesValue = "Trim(($SpeciesField & ''))"
    return "($SpeciesField Is Null OR Len($speciesValue) = 0)"
}

function Get-RequiredTreeCrownHistoryCodes {
    return @("0", "5", "10")
}

function Get-TreeCrownRequiredCondition {
    param(
        [string]$TreeHistoryField = "[TreeHistory]",
        [string]$CrownField = "[CrownRatio]"
    )

    $historyValue = "Trim(($TreeHistoryField & ''))"
    $historyList = Convert-CodeValuesToSqlList -Values (Get-RequiredTreeCrownHistoryCodes)
    return "($historyValue In ($historyList) AND ($CrownField Is Null OR Len(Trim(($CrownField & ''))) = 0))"
}

function Get-TreeRadialIncrementRequiredCondition {
    param(
        [string]$TreeHistoryField = "tm.[TreeHistory]",
        [string]$RadialIncrementField = "tm.[RadialIncrement]",
        [string]$SpeciesField = "",
        [string[]]$ProblemFields = @()
    )

    $historyValue = "Trim(($TreeHistoryField & ''))"
    $blankRadial = "($RadialIncrementField Is Null OR Len(Trim(($RadialIncrementField & ''))) = 0)"
    $historyFiveRequired = "$historyValue = '5'"
    $problem121Required = ""
    $problem121Parts = New-Object System.Collections.Generic.List[string]
    foreach ($problemField in @($ProblemFields)) {
        if ([string]::IsNullOrWhiteSpace($problemField)) { continue }
        [void]$problem121Parts.Add("Trim(($problemField & '')) = '121'")
    }
    if ($problem121Parts.Count -gt 0) {
        $historyList = Convert-CodeValuesToSqlList -Values @("0", "5", "10")
        $problem121Condition = "(" + ([string]::Join(" OR ", [string[]]$problem121Parts.ToArray())) + ")"
        $problem121Required = "($historyValue In ($historyList) AND $problem121Condition)"
    }

    $requiredReason = $historyFiveRequired
    if (-not [string]::IsNullOrWhiteSpace($problem121Required)) {
        $requiredReason = "(($historyFiveRequired) OR ($problem121Required))"
    }

    $parts = New-Object System.Collections.Generic.List[string]
    [void]$parts.Add($requiredReason)
    [void]$parts.Add($blankRadial)

    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    if (-not [string]::IsNullOrWhiteSpace($SpeciesField) -and -not [string]::IsNullOrWhiteSpace($woodlandList)) {
        [void]$parts.Add("(Trim(($SpeciesField & '')) = '' OR Trim(($SpeciesField & '')) Not In ($woodlandList))")
    }

    return "(" + ([string]::Join(" AND ", [string[]]$parts.ToArray())) + ")"
}

function Get-RequiredTreeHeightHistoryCodes {
    $codes = New-Object System.Collections.Generic.List[string]
    $mode = Get-TotalHeightProtocolMode

    if ($mode -eq "FullLive") {
        foreach ($code in @("0", "5", "10")) {
            Add-UniqueCode -Codes $codes -Code $code
        }
    }
    elseif ($mode -ne "Subsample") {
        foreach ($code in @("0", "10")) {
            Add-UniqueCode -Codes $codes -Code $code
        }
    }

    if (Test-NewMortalityHeightChecksEnabled) {
        foreach ($code in @("2", "3")) {
            Add-UniqueCode -Codes $codes -Code $code
        }
    }

    if (Test-OldMortalityHeightChecksEnabled) {
        Add-UniqueCode -Codes $codes -Code "7"
    }

    return [string[]]$codes.ToArray()
}

function Get-TreeHeightRequiredHistoryText {
    $codes = @(Get-RequiredTreeHeightHistoryCodes)
    if ($codes.Count -eq 0) { return "no TreeHistory statuses" }
    return "TreeHistory " + ([string]::Join(", ", [string[]]$codes))
}

function Get-DeadTreeHeightShrinkAllowedHistoryCodes {
    $codes = New-Object System.Collections.Generic.List[string]
    if (Test-NewMortalityHeightChecksEnabled) {
        foreach ($code in @("2", "3")) {
            Add-UniqueCode -Codes $codes -Code $code
        }
    }

    if (Test-OldMortalityHeightChecksEnabled) {
        Add-UniqueCode -Codes $codes -Code "7"
    }

    return [string[]]$codes.ToArray()
}

function Get-DeadTreeHeightShrinkAllowedCondition {
    param(
        [string]$EarlierAlias = "earlier",
        [string]$LaterAlias = "later"
    )

    $deadCodes = @(Get-DeadTreeHeightShrinkAllowedHistoryCodes)
    if ($deadCodes.Count -eq 0) { return "" }

    $deadCodeList = [string]::Join(", ", [string[]]$deadCodes)
    $earlierHistoryText = "Trim(($EarlierAlias.[TreeHistory] & ''))"
    $laterHistoryText = "Trim(($LaterAlias.[TreeHistory] & ''))"

    return "($earlierHistoryText <> '' AND Val($earlierHistoryText) In (0, 5, 10) AND $laterHistoryText <> '' AND Val($laterHistoryText) In ($deadCodeList) AND $LaterAlias.[TotalHeight] < $EarlierAlias.[TotalHeight])"
}

function Get-HeightRequiredMinorPlotValues {
    $rawText = Get-ControlTextValue -Control $heightRequiredMinorPlotsBox -DefaultValue $script:DefaultHeightRequiredMinorPlots -SettingName "HeightRequiredMinorPlotsText"
    return @(Convert-DelimitedCodeTextToValues -Text $rawText)
}

function Get-HeightRequiredMinorPlotsSqlList {
    $codes = @(Get-HeightRequiredMinorPlotValues)
    return (Convert-CodeValuesToSqlList -Values $codes)
}

function Get-TreeHeightMissingRequiredCondition {
    param(
        [string]$TreeHistoryField = "[TreeHistory]",
        [string]$HeightField = "[TotalHeight]",
        [string]$MinorPlotExpression = "",
        [object[]]$MeasurementColumns = @(),
        [object[]]$TreeColumns = @(),
        [string]$MeasurementAlias = "",
        [string]$TreeAlias = ""
    )

    $historyCodes = @(Get-RequiredTreeHeightHistoryCodes)
    if ($historyCodes.Count -eq 0) { return "(1 = 0)" }

    if (-not [string]::IsNullOrWhiteSpace($MeasurementAlias)) {
        $historyRequired = Get-TreeHeightStatusSqlCondition `
            -MeasurementColumns $MeasurementColumns `
            -TreeColumns $TreeColumns `
            -MeasurementAlias $MeasurementAlias `
            -TreeAlias $TreeAlias `
            -TreeHistoryCodes $historyCodes
    }
    else {
        $historyValue = "Trim(($TreeHistoryField & ''))"
        $historyList = ([string[]]@($historyCodes | ForEach-Object { Sql-Text $_ })) -join ", "
        $historyRequired = "($historyValue In ($historyList))"
    }

    $condition = "(($historyRequired) AND $HeightField Is Null)"

    $minorPlotList = Get-HeightRequiredMinorPlotsSqlList
    if ((-not (Test-TotalHeightFullLiveProtocolEnabled)) -and -not [string]::IsNullOrWhiteSpace($minorPlotList)) {
        if ([string]::IsNullOrWhiteSpace($MinorPlotExpression)) { return "" }
        $condition = "($condition AND $MinorPlotExpression In ($minorPlotList))"
    }

    return $condition
}

function Get-TreeHeightRareSpeciesMissingCondition {
    param(
        [object[]]$MeasurementColumns,
        [object[]]$TreeColumns,
        [string]$MeasurementAlias = "tm",
        [string]$TreeAlias = "treeMeta",
        [string]$HeightField = "tm.[TotalHeight]"
    )

    $rareSpeciesList = Get-HeightRareSpeciesCodesSqlList
    if ([string]::IsNullOrWhiteSpace($rareSpeciesList)) { return "" }
    if ([string]::IsNullOrWhiteSpace($TreeAlias) -or -not (Test-ColumnExists -Columns $TreeColumns -Name "SpeciesCode")) { return "" }

    $liveTreeCondition = Get-TreeHeightStatusSqlCondition `
        -MeasurementColumns $MeasurementColumns `
        -TreeColumns $TreeColumns `
        -MeasurementAlias $MeasurementAlias `
        -TreeAlias $TreeAlias `
        -TreeHistoryCodes @("0", "5", "10")
    $speciesValue = "Trim(($TreeAlias.[SpeciesCode] & ''))"
    return "(($liveTreeCondition) AND $speciesValue In ($rareSpeciesList) AND $HeightField Is Null)"
}

function Get-ProblemCodeHeightExceptionCondition {
    param(
        [object[]]$MeasurementColumns,
        [object[]]$TreeColumns,
        [string]$MeasurementAlias = "tm",
        [string]$TreeAlias = "treeMeta"
    )

    $problemCodes = @(Get-ProblemHeightNoHeightCodes)
    if ($problemCodes.Count -eq 0) { return "" }
    if ([string]::IsNullOrWhiteSpace($TreeAlias) -or -not (Test-ColumnExists -Columns $TreeColumns -Name "SpeciesCode")) { return "" }

    $problemParts = New-Object System.Collections.Generic.List[string]
    $problemList = Convert-CodeValuesToSqlList -Values $problemCodes
    foreach ($problemField in @("Problem1", "Problem2")) {
        if (-not (Test-ColumnExists -Columns $MeasurementColumns -Name $problemField)) { continue }
        [void]$problemParts.Add("Trim(($MeasurementAlias.$(Quote-Name $problemField) & '')) In ($problemList)")
    }
    if ($problemParts.Count -eq 0) { return "" }

    $speciesValue = "Trim(($TreeAlias.[SpeciesCode] & ''))"
    $timberCondition = "$speciesValue <> ''"
    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    if (-not [string]::IsNullOrWhiteSpace($woodlandList)) {
        $timberCondition = "($timberCondition AND $speciesValue Not In ($woodlandList))"
    }

    $problemCondition = [string]::Join(" OR ", [string[]]$problemParts.ToArray())
    return "(($problemCondition) AND ($timberCondition))"
}

function Get-ProblemCodeHeightShouldBeBlankCondition {
    param(
        [object[]]$MeasurementColumns,
        [object[]]$TreeColumns,
        [string]$MeasurementAlias = "tm",
        [string]$TreeAlias = "treeMeta",
        [string]$HeightField = "tm.[TotalHeight]"
    )

    $problemCodes = @(Get-ProblemHeightNoHeightCodes)
    if ($problemCodes.Count -eq 0) { return "" }
    if ([string]::IsNullOrWhiteSpace($TreeAlias) -or -not (Test-ColumnExists -Columns $TreeColumns -Name "SpeciesCode")) { return "" }

    $problemParts = New-Object System.Collections.Generic.List[string]
    $problemList = Convert-CodeValuesToSqlList -Values $problemCodes
    foreach ($problemField in @("Problem1", "Problem2")) {
        if (-not (Test-ColumnExists -Columns $MeasurementColumns -Name $problemField)) { continue }
        [void]$problemParts.Add("Trim(($MeasurementAlias.$(Quote-Name $problemField) & '')) In ($problemList)")
    }
    if ($problemParts.Count -eq 0) { return "" }

    $speciesValue = "Trim(($TreeAlias.[SpeciesCode] & ''))"
    $timberCondition = "$speciesValue <> ''"
    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    if (-not [string]::IsNullOrWhiteSpace($woodlandList)) {
        $timberCondition = "($timberCondition AND $speciesValue Not In ($woodlandList))"
    }

    $problemCondition = [string]::Join(" OR ", [string[]]$problemParts.ToArray())
    return "($HeightField Is Not Null AND ($problemCondition) AND ($timberCondition))"
}

function Get-ProblemCodeHeightBlankAllowedCondition {
    param(
        [object[]]$MeasurementColumns,
        [object[]]$TreeColumns,
        [string]$MeasurementAlias = "tm",
        [string]$TreeAlias = "treeMeta",
        [string]$HeightField = "tm.[TotalHeight]"
    )

    $problemCodes = @(Get-ProblemHeightNoHeightCodes)
    if ($problemCodes.Count -eq 0) { return "" }
    if ([string]::IsNullOrWhiteSpace($TreeAlias) -or -not (Test-ColumnExists -Columns $TreeColumns -Name "SpeciesCode")) { return "" }

    $problemParts = New-Object System.Collections.Generic.List[string]
    $problemList = Convert-CodeValuesToSqlList -Values $problemCodes
    foreach ($problemField in @("Problem1", "Problem2")) {
        if (-not (Test-ColumnExists -Columns $MeasurementColumns -Name $problemField)) { continue }
        [void]$problemParts.Add("Trim(($MeasurementAlias.$(Quote-Name $problemField) & '')) In ($problemList)")
    }
    if ($problemParts.Count -eq 0) { return "" }

    $speciesValue = "Trim(($TreeAlias.[SpeciesCode] & ''))"
    $timberCondition = "$speciesValue <> ''"
    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    if (-not [string]::IsNullOrWhiteSpace($woodlandList)) {
        $timberCondition = "($timberCondition AND $speciesValue Not In ($woodlandList))"
    }

    $problemCondition = [string]::Join(" OR ", [string[]]$problemParts.ToArray())
    return "($HeightField Is Null AND ($problemCondition) AND ($timberCondition))"
}

function Assert-GuideRunNotCanceled {
    $cancelRequested = [bool]$script:GuideRunCancellationRequested
    try {
        if ($null -ne $script:GuideRunWorker -and $script:GuideRunWorker.CancellationPending) {
            $cancelRequested = $true
        }
    }
    catch {
    }

    try {
        if (-not [string]::IsNullOrWhiteSpace($script:GuideRunCancelPath) -and (Test-Path -LiteralPath $script:GuideRunCancelPath)) {
            $cancelRequested = $true
        }
    }
    catch {
    }

    if ($cancelRequested) {
        throw [System.OperationCanceledException]::new("Run canceled.")
    }
}

function Get-WoodlandSpeciesCodeValues {
    $rawText = Get-ControlTextValue -Control $woodlandSpeciesCodesBox -DefaultValue $script:DefaultWoodlandSpeciesCodes -SettingName "WoodlandSpeciesCodesText"
    return @(Convert-DelimitedCodeTextToValues -Text $rawText)
}

function Convert-DelimitedCodeTextToValues {
    param([string]$Text)

    $rawText = [string]$Text
    if ([string]::IsNullOrWhiteSpace($rawText)) { return @() }

    $codes = New-Object System.Collections.Generic.List[string]
    foreach ($part in @($rawText -split "[,;|\s]+")) {
        $code = $part.Trim()
        if ([string]::IsNullOrWhiteSpace($code)) { continue }
        Add-UniqueCode -Codes $codes -Code $code
        if ($code -match "^\d+$") {
            Add-UniqueCode -Codes $codes -Code (([int64]$code).ToString())
        }
    }

    return @($codes.ToArray())
}

function Convert-CodeValuesToSqlList {
    param([string[]]$Values)

    $codes = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
    if ($codes.Count -eq 0) { return "" }
    return (($codes | ForEach-Object { Sql-Text $_ }) -join ", ")
}

function Get-WoodlandSpeciesCodesSqlList {
    $codes = @(Get-WoodlandSpeciesCodeValues)
    return (Convert-CodeValuesToSqlList -Values $codes)
}

function Get-RegenTimberSeedlingMinorPlotValues {
    return @(Convert-DelimitedCodeTextToValues -Text (Get-ControlTextValue -Control $regenTimberSeedlingMinorPlotsBox -DefaultValue $script:DefaultRegenTimberSeedlingMinorPlots -SettingName "RegenTimberSeedlingMinorPlotsText"))
}

function Get-RegenTimberSapling20MinorPlotValues {
    return @(Convert-DelimitedCodeTextToValues -Text (Get-ControlTextValue -Control $regenTimberSapling20MinorPlotsBox -DefaultValue $script:DefaultRegenTimberSapling20MinorPlots -SettingName "RegenTimberSapling20MinorPlotsText"))
}

function Get-RegenTimberSapling40MinorPlotValues {
    return @(Convert-DelimitedCodeTextToValues -Text (Get-ControlTextValue -Control $regenTimberSapling40MinorPlotsBox -DefaultValue $script:DefaultRegenTimberSapling40MinorPlots -SettingName "RegenTimberSapling40MinorPlotsText"))
}

function Get-RegenWoodlandSeedlingMinorPlotValues {
    return @(Convert-DelimitedCodeTextToValues -Text (Get-ControlTextValue -Control $regenWoodlandSeedlingMinorPlotsBox -DefaultValue $script:DefaultRegenWoodlandSeedlingMinorPlots -SettingName "RegenWoodlandSeedlingMinorPlotsText"))
}

function Get-RegenWoodlandSapling20MinorPlotValues {
    return @(Convert-DelimitedCodeTextToValues -Text (Get-ControlTextValue -Control $regenWoodlandSapling20MinorPlotsBox -DefaultValue $script:DefaultRegenWoodlandSapling20MinorPlots -SettingName "RegenWoodlandSapling20MinorPlotsText"))
}

function Get-RegenWoodlandSapling40MinorPlotValues {
    return @(Convert-DelimitedCodeTextToValues -Text (Get-ControlTextValue -Control $regenWoodlandSapling40MinorPlotsBox -DefaultValue $script:DefaultRegenWoodlandSapling40MinorPlots -SettingName "RegenWoodlandSapling40MinorPlotsText"))
}

function Get-WoodlandSpeciesExclusionSql {
    param([string]$TreeAlias)

    if ([string]::IsNullOrWhiteSpace($TreeAlias)) { return "" }

    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    if ([string]::IsNullOrWhiteSpace($woodlandList)) { return "" }

    return "(Trim(($TreeAlias.[SpeciesCode] & '')) = '' OR Trim(($TreeAlias.[SpeciesCode] & '')) Not In ($woodlandList))"
}

function Get-LiveTimberDbhShrinkSqlCondition {
    param(
        [string]$CurrentAlias,
        [string]$PreviousAlias,
        [string]$TreeAlias,
        [string]$PlotAlias
    )

    $parts = New-Object System.Collections.Generic.List[string]
    [void]$parts.Add("$CurrentAlias.[TreeHistory] In (0, 10)")
    [void]$parts.Add("$CurrentAlias.[IDBH] < $PreviousAlias.[IDBH]")

    $woodlandCondition = Get-WoodlandSpeciesExclusionSql -TreeAlias $TreeAlias
    if (-not [string]::IsNullOrWhiteSpace($woodlandCondition)) {
        [void]$parts.Add($woodlandCondition)
    }

    if (-not [string]::IsNullOrWhiteSpace($PlotAlias)) {
        [void]$parts.Add("($PlotAlias.[PlotKindID] Is Null OR $PlotAlias.[PlotKindID] = 1)")
    }

    return "(" + ([string]::Join(" AND ", [string[]]$parts.ToArray())) + ")"
}

function Get-WoodlandShrinkSkipVerification {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $enteredCodes = @(Get-WoodlandSpeciesCodeValues | Sort-Object -Unique)
    $enteredText = if ($enteredCodes.Count -gt 0) { [string]::Join(", ", [string[]]$enteredCodes) } else { "" }

    if ($enteredCodes.Count -eq 0) {
        return [pscustomobject]@{
            Status = "No codes entered"
            EnteredCodes = ""
            SkippedRows = 0
            FlaggedRows = 0
            MatchedCodes = ""
            Message = "No woodland species codes were entered, so no species were excluded from shrinking diameter checks."
        }
    }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) {
        return [pscustomobject]@{
            Status = "Not checked"
            EnteredCodes = $enteredText
            SkippedRows = 0
            FlaggedRows = 0
            MatchedCodes = ""
            Message = "Woodland species were entered ($enteredText), but TreeMeasurements was not found, so the shrinking diameter skip could not be verified."
        }
    }

    $columns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredField in @("MeasurementID", "TreeID", "PeriodNumber", "TreeHistory", "IDBH")) {
        if (-not (Test-ColumnExists -Columns $columns -Name $requiredField)) {
            return [pscustomobject]@{
                Status = "Not checked"
                EnteredCodes = $enteredText
                SkippedRows = 0
                FlaggedRows = 0
                MatchedCodes = ""
                Message = "Woodland species were entered ($enteredText), but TreeMeasurements.$requiredField was not found, so the shrinking diameter skip could not be verified."
            }
        }
    }

    $metadataAvailability = Get-TreeMetadataAvailability -Connection $Connection -Tables $tables -MeasurementColumns $columns
    if (-not $metadataAvailability.IncludeTreeMetadata) {
        return [pscustomobject]@{
            Status = "Not applied"
            EnteredCodes = $enteredText
            SkippedRows = 0
            FlaggedRows = 0
            MatchedCodes = ""
            Message = "Woodland species were entered ($enteredText), but Trees.SpeciesCode could not be joined to TreeMeasurements. The species skip could not be applied or verified."
        }
    }

    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    if ([string]::IsNullOrWhiteSpace($woodlandList)) {
        return [pscustomobject]@{
            Status = "No usable codes"
            EnteredCodes = $enteredText
            SkippedRows = 0
            FlaggedRows = 0
            MatchedCodes = ""
            Message = "Woodland species text was entered, but no usable species codes were parsed."
        }
    }

    $fromSql = Get-TreeMeasurementPairFromSql `
        -IncludeTreeMetadata `
        -IncludePlotMetadata:($metadataAvailability.IncludePlotMetadata)

    $conditions = New-Object System.Collections.Generic.List[string]
    [void]$conditions.Add("earlier.[IDBH] Is Not Null")
    [void]$conditions.Add("later.[IDBH] Is Not Null")
    [void]$conditions.Add("later.[PeriodNumber] > earlier.[PeriodNumber]")
    $laterPeriodScope = Get-SelectedPeriodAliasCondition -Alias "later"
    if (-not [string]::IsNullOrWhiteSpace($laterPeriodScope)) {
        [void]$conditions.Add($laterPeriodScope)
    }
    [void]$conditions.Add("later.[TreeHistory] In (0, 10)")
    [void]$conditions.Add("later.[IDBH] < earlier.[IDBH]")
    [void]$conditions.Add("Trim((treeMeta.[SpeciesCode] & '')) In ($woodlandList)")
    if ($metadataAvailability.IncludePlotMetadata) {
        [void]$conditions.Add("(plotMeta.[PlotKindID] Is Null OR plotMeta.[PlotKindID] = 1)")
    }
    $whereClause = [string]::Join(" AND ", [string[]]$conditions.ToArray())

    try {
        $skippedRows = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) $fromSql WHERE $whereClause"
        $matchedTable = Get-DataTable -Connection $Connection -Sql "SELECT DISTINCT Trim((treeMeta.[SpeciesCode] & '')) AS [SpeciesCode] $fromSql WHERE $whereClause"
        $matchedCodes = New-Object System.Collections.Generic.List[string]
        foreach ($row in $matchedTable.Rows) {
            Add-UniqueCode -Codes $matchedCodes -Code (Normalize-CodeValue $row["SpeciesCode"])
        }
        $matchedText = if ($matchedCodes.Count -gt 0) { [string]::Join(", ", [string[]]$matchedCodes.ToArray()) } else { "none found with shrinking DBH" }
        $flaggedRows = 0
        if (Test-TableAvailable -Tables $tables -TableName "InventoryCleanAudit") {
            $flaggedRows += Get-CountValue -Connection $Connection -Sql @"
SELECT Count(*)
FROM ([InventoryCleanAudit] AS a
INNER JOIN [TreeMeasurements] AS later
    ON a.[SourceRowId] = (later.[MeasurementID] & ''))
LEFT JOIN [Trees] AS treeMeta
    ON later.[TreeID] = treeMeta.[TreeID]
WHERE a.[TableName] = 'TreeMeasurements'
  AND a.[RuleName] = 'IDBH shrinkage check'
  AND Trim((treeMeta.[SpeciesCode] & '')) In ($woodlandList)
"@

            $previousFromSql = Get-TreePreviousMeasurementFromSql `
                -IncludeTreeMetadata `
                -IncludePlotMetadata:($metadataAvailability.IncludePlotMetadata)
            $previousConditions = New-Object System.Collections.Generic.List[string]
            [void]$previousConditions.Add("cur.[PeriodNumber] Is Not Null")
            $currentPeriodScope = Get-SelectedPeriodAliasCondition -Alias "cur"
            if (-not [string]::IsNullOrWhiteSpace($currentPeriodScope)) {
                [void]$previousConditions.Add($currentPeriodScope)
            }
            [void]$previousConditions.Add("cur.[IDBH] Is Not Null")
            [void]$previousConditions.Add("prev.[IDBH] Is Not Null")
            [void]$previousConditions.Add("cur.[TreeHistory] In (0, 10)")
            [void]$previousConditions.Add("cur.[IDBH] < prev.[IDBH]")
            [void]$previousConditions.Add("Trim((treeMeta.[SpeciesCode] & '')) In ($woodlandList)")
            if ($metadataAvailability.IncludePlotMetadata) {
                [void]$previousConditions.Add("(plotMeta.[PlotKindID] Is Null OR plotMeta.[PlotKindID] = 1)")
            }
            $previousWhereClause = [string]::Join(" AND ", [string[]]$previousConditions.ToArray())
            $flaggedRows += Get-CountValue -Connection $Connection -Sql @"
SELECT Count(*)
FROM [InventoryCleanAudit] AS a
INNER JOIN (
    SELECT cur.[MeasurementID] AS [MeasurementID]
    $previousFromSql
    WHERE $previousWhereClause
) AS skipped
    ON a.[SourceRowId] = (skipped.[MeasurementID] & '')
WHERE a.[TableName] = 'TreeMeasurements'
  AND a.[RuleName] = 'IDBH jump check'
"@
        }
        $status = if ($flaggedRows -eq 0) { "Verified" } else { "Review" }
        $message = "Woodland species excluded from shrinking diameter checks: $enteredText. Skipped shrinking DBH pair(s): $skippedRows. Matched skipped species: $matchedText. Shrinking DBH findings left for those species: $flaggedRows."

        return [pscustomobject]@{
            Status = $status
            EnteredCodes = $enteredText
            SkippedRows = $skippedRows
            FlaggedRows = $flaggedRows
            MatchedCodes = $matchedText
            Message = $message
        }
    }
    catch {
        return [pscustomobject]@{
            Status = "Not checked"
            EnteredCodes = $enteredText
            SkippedRows = 0
            FlaggedRows = 0
            MatchedCodes = ""
            Message = "Woodland species were entered ($enteredText), but skip verification failed: $($_.Exception.Message)"
        }
    }
}

function Get-TreeMetadataAvailability {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string[]]$Tables,
        [object[]]$MeasurementColumns
    )

    $includeTreeMetadata = $false
    $includePlotMetadata = $false

    if ((Test-TableAvailable -Tables $Tables -TableName "Trees") -and
        (Test-ColumnExists -Columns $MeasurementColumns -Name "TreeID")) {

        $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
        $includeTreeMetadata =
            (Test-ColumnExists -Columns $treeColumns -Name "TreeID") -and
            (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")

        if ($includeTreeMetadata -and
            (Test-TableAvailable -Tables $Tables -TableName "Plots") -and
            (Test-ColumnExists -Columns $treeColumns -Name "PlotID")) {

            $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
            $includePlotMetadata =
                (Test-ColumnExists -Columns $plotColumns -Name "PlotID") -and
                (Test-ColumnExists -Columns $plotColumns -Name "PlotKindID")
        }
    }

    return [pscustomobject]@{
        IncludeTreeMetadata = $includeTreeMetadata
        IncludePlotMetadata = $includePlotMetadata
    }
}

function Get-TreeHistoryClassCrosswalkSummary {
    return "TreeClass/TreeStatus to TreeHistory crosswalk reference from TreeClass_Status_History_Crosswalk.xlsx: TreeClass 1/2/3/4/9 are treated as live and usually TreeHistory 0, 5, or 10; class 5 maps to harvest TreeHistory 1; class 6 maps to thinned TreeHistory 4; class 7 maps to new salvageable mortality TreeHistory 2; class 8 maps to new non-salvageable mortality TreeHistory 3; class 0 is conditional and needs problem-code, previous-history, or project-use review. TreeStatus 1/2/3 are treated as live and usually TreeHistory 0, 5, or 10; status 5 is harvest, status 4 is conditional current mortality/thin review, and status 9 is previous harvest/mortality review."
}

function Get-TreeHistoryConversionCandidateNames {
    param([string]$SourceKind)

    switch ($SourceKind) {
        "TreeClass" {
            return @(
                "TreeClass",
                "TreeClassCode",
                "TreeClassID",
                "TreeClassCd",
                "TreeClassCdID"
            )
        }
        "TreeStatus" {
            return @(
                "TreeStatus",
                "TreeStatusCode",
                "TreeStatusID",
                "TreeStatusCd",
                "TreeStatusCdID"
            )
        }
    }

    return @()
}

function Get-LiveTreeClassCodes {
    return @("1", "2", "3", "4", "9")
}

function Get-LiveTreeStatusCodes {
    return @("1", "2", "3")
}

function Test-TreeHistoryCodeListIncludesLive {
    param([string[]]$Codes)

    foreach ($code in @($Codes)) {
        if ($code -in @("0", "5", "10")) { return $true }
    }

    return $false
}

function Add-LegacyLiveSourceConditionParts {
    param(
        [System.Collections.Generic.List[string]]$Parts,
        [object[]]$MeasurementColumns,
        [object[]]$TreeColumns,
        [string]$MeasurementAlias,
        [string]$TreeAlias
    )

    foreach ($sourceKind in @("TreeClass", "TreeStatus")) {
        $sourceCodes = if ($sourceKind -eq "TreeClass") { @(Get-LiveTreeClassCodes) } else { @(Get-LiveTreeStatusCodes) }
        $sourceList = Convert-CodeValuesToSqlList -Values $sourceCodes
        if ([string]::IsNullOrWhiteSpace($sourceList)) { continue }

        if (-not [string]::IsNullOrWhiteSpace($MeasurementAlias)) {
            $measurementField = Find-ExactCandidateColumn -Columns $MeasurementColumns -CandidateNames (Get-TreeHistoryConversionCandidateNames -SourceKind $sourceKind)
            if (-not [string]::IsNullOrWhiteSpace($measurementField)) {
                [void]$Parts.Add("Trim(($MeasurementAlias.$(Quote-Name $measurementField) & '')) In ($sourceList)")
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($TreeAlias)) {
            $treeField = Find-ExactCandidateColumn -Columns $TreeColumns -CandidateNames (Get-TreeHistoryConversionCandidateNames -SourceKind $sourceKind)
            if (-not [string]::IsNullOrWhiteSpace($treeField)) {
                [void]$Parts.Add("Trim(($TreeAlias.$(Quote-Name $treeField) & '')) In ($sourceList)")
            }
        }
    }
}

function Get-TreeHeightStatusSqlCondition {
    param(
        [object[]]$MeasurementColumns,
        [object[]]$TreeColumns,
        [string]$MeasurementAlias = "tm",
        [string]$TreeAlias = "",
        [string[]]$TreeHistoryCodes = @("0", "5", "10")
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $historyList = Convert-CodeValuesToSqlList -Values $TreeHistoryCodes
    if (-not [string]::IsNullOrWhiteSpace($MeasurementAlias) -and
        -not [string]::IsNullOrWhiteSpace($historyList) -and
        (Test-ColumnExists -Columns $MeasurementColumns -Name "TreeHistory")) {
        [void]$parts.Add("Trim(($MeasurementAlias.[TreeHistory] & '')) In ($historyList)")
    }

    if (Test-TreeHistoryCodeListIncludesLive -Codes $TreeHistoryCodes) {
        Add-LegacyLiveSourceConditionParts `
            -Parts $parts `
            -MeasurementColumns $MeasurementColumns `
            -TreeColumns $TreeColumns `
            -MeasurementAlias $MeasurementAlias `
            -TreeAlias $TreeAlias
    }

    if ($parts.Count -eq 0) { return "(1 = 0)" }
    return "((" + ([string]::Join(") OR (", [string[]]$parts.ToArray())) + "))"
}

function Find-ExactCandidateColumn {
    param(
        [object[]]$Columns,
        [string[]]$CandidateNames
    )

    foreach ($candidate in @($CandidateNames)) {
        $candidateNormalized = Normalize-FieldName $candidate
        if ([string]::IsNullOrWhiteSpace($candidateNormalized)) { continue }
        foreach ($column in @($Columns)) {
            if ((Normalize-FieldName ([string]$column.Name)) -eq $candidateNormalized) {
                return [string]$column.Name
            }
        }
    }

    return $null
}

function Get-TreeHistoryConversionSourceDefinition {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$SourceKind
    )

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return $null }

    $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredField in @("MeasurementID", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name $requiredField)) { return $null }
    }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "TreeHistory")) { return $null }

    $candidateNames = @(Get-TreeHistoryConversionCandidateNames -SourceKind $SourceKind)
    if ($candidateNames.Count -eq 0) { return $null }

    $sourceField = Find-ExactCandidateColumn -Columns $treeMeasurementColumns -CandidateNames $candidateNames
    if (-not [string]::IsNullOrWhiteSpace($sourceField) -and
        (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName $sourceField)) {
        return [pscustomobject]@{
            SourceKind = $SourceKind
            SourceTable = "TreeMeasurements"
            SourceField = $sourceField
            SourceAlias = "tm"
            SourceExpression = "Trim((tm.$(Quote-Name $sourceField) & ''))"
            FromSql = "FROM [TreeMeasurements] AS tm"
            RecordLabelExpression = (Get-TreeStemCountRecordLabelExpression -Columns $treeMeasurementColumns -Alias "tm")
        }
    }

    if (-not (Test-TableAvailable -Tables $tables -TableName "Trees")) { return $null }
    if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name "TreeID")) { return $null }

    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "TreeID")) { return $null }

    $sourceField = Find-ExactCandidateColumn -Columns $treeColumns -CandidateNames $candidateNames
    if ([string]::IsNullOrWhiteSpace($sourceField)) { return $null }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "Trees" -FieldName $sourceField)) { return $null }

    return [pscustomobject]@{
        SourceKind = $SourceKind
        SourceTable = "Trees"
        SourceField = $sourceField
        SourceAlias = "treeSource"
        SourceExpression = "Trim((treeSource.$(Quote-Name $sourceField) & ''))"
        FromSql = @"
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeSource
    ON tm.[TreeID] = treeSource.[TreeID]
"@
        RecordLabelExpression = (Get-TreeStemCountRecordLabelExpression -Columns $treeMeasurementColumns -Alias "tm")
    }
}

function Get-TreeClassHistoryMismatchCondition {
    param(
        [string]$SourceExpression,
        [string]$HistoryExpression
    )

    return @"
(Len($SourceExpression) > 0 AND Len($HistoryExpression) > 0 AND (
        ($SourceExpression In ('1', '2', '3', '4', '9') AND $HistoryExpression Not In ('0', '5', '10'))
        OR ($SourceExpression = '5' AND $HistoryExpression <> '1')
        OR ($SourceExpression = '6' AND $HistoryExpression <> '4')
        OR ($SourceExpression = '7' AND $HistoryExpression <> '2')
        OR ($SourceExpression = '8' AND $HistoryExpression <> '3')
        OR ($SourceExpression = '0' AND $HistoryExpression Not In ('5', '6', '7', '8'))
    ))
"@
}

function Get-TreeClassHistoryReviewCondition {
    param(
        [string]$SourceExpression,
        [string]$HistoryExpression
    )

    return @"
(Len($SourceExpression) > 0 AND Len($HistoryExpression) > 0 AND (
        ($SourceExpression = '0' AND $HistoryExpression In ('5', '6', '7', '8'))
        OR ($SourceExpression = '7' AND $HistoryExpression = '2')
        OR ($SourceExpression = '8' AND $HistoryExpression = '3')
        OR ($SourceExpression = '10')
    ))
"@
}

function Get-TreeStatusHistoryMismatchCondition {
    param(
        [string]$SourceExpression,
        [string]$HistoryExpression
    )

    return @"
(Len($SourceExpression) > 0 AND Len($HistoryExpression) > 0 AND (
        ($SourceExpression In ('1', '2', '3') AND $HistoryExpression Not In ('0', '5', '10'))
        OR ($SourceExpression = '5' AND $HistoryExpression <> '1')
        OR ($SourceExpression = '4' AND $HistoryExpression Not In ('2', '3', '4'))
        OR ($SourceExpression = '9' AND $HistoryExpression Not In ('7', '8'))
    ))
"@
}

function Get-TreeStatusHistoryReviewCondition {
    param(
        [string]$SourceExpression,
        [string]$HistoryExpression
    )

    return @"
(Len($SourceExpression) > 0 AND Len($HistoryExpression) > 0 AND (
        ($SourceExpression = '4' AND $HistoryExpression In ('2', '3', '4'))
        OR ($SourceExpression = '9' AND $HistoryExpression In ('7', '8'))
    ))
"@
}

function Get-TreeHistoryConversionCondition {
    param(
        [string]$SourceKind,
        [string]$Mode,
        [string]$SourceExpression,
        [string]$HistoryExpression
    )

    if ($SourceKind -eq "TreeClass") {
        if ($Mode -eq "Review") {
            return Get-TreeClassHistoryReviewCondition -SourceExpression $SourceExpression -HistoryExpression $HistoryExpression
        }
        return Get-TreeClassHistoryMismatchCondition -SourceExpression $SourceExpression -HistoryExpression $HistoryExpression
    }

    if ($SourceKind -eq "TreeStatus") {
        if ($Mode -eq "Review") {
            return Get-TreeStatusHistoryReviewCondition -SourceExpression $SourceExpression -HistoryExpression $HistoryExpression
        }
        return Get-TreeStatusHistoryMismatchCondition -SourceExpression $SourceExpression -HistoryExpression $HistoryExpression
    }

    return ""
}

function Get-TreeHistoryConversionRuleName {
    param(
        [string]$SourceKind,
        [string]$Mode
    )

    if ($Mode -eq "Review") { return "$SourceKind/TreeHistory conversion review" }
    return "$SourceKind/TreeHistory conversion mismatch"
}

function Get-TreeHistoryConversionFieldName {
    param([string]$SourceKind)
    return "$SourceKind vs TreeHistory"
}

function Get-TreeHistoryConversionMessage {
    param(
        [string]$SourceKind,
        [string]$Mode,
        [string]$SourceTable,
        [string]$SourceField
    )

    $sourceText = "$SourceTable.$SourceField"
    if ($SourceKind -eq "TreeClass") {
        if ($Mode -eq "Review") {
            return "TreeClass-to-TreeHistory conversion uses a conditional crosswalk case from $sourceText. Review problem codes, previous TreeHistory, snag/downed status, and project use before changing TreeHistory. " + (Get-TreeHistoryClassCrosswalkSummary)
        }
        return "TreeClass-to-TreeHistory conversion does not match the clear crosswalk mapping from $sourceText. Review the source class and TreeHistory before editing. " + (Get-TreeHistoryClassCrosswalkSummary)
    }

    if ($Mode -eq "Review") {
        return "TreeStatus-to-TreeHistory conversion uses a conditional crosswalk case from $sourceText. Review snag/downed status, previous TreeHistory, and project use before changing TreeHistory. " + (Get-TreeHistoryClassCrosswalkSummary)
    }
    return "TreeStatus-to-TreeHistory conversion does not match the clear crosswalk mapping from $sourceText. Review the source status and TreeHistory before editing. " + (Get-TreeHistoryClassCrosswalkSummary)
}

function Add-TreeHistoryConversionChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return 0 }

    Ensure-NeedsReviewColumn -Connection $Connection -TableName "TreeMeasurements"
    $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredField in @("MeasurementID", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name $requiredField)) { return 0 }
    }

    $historyExpression = "Trim((tm.[TreeHistory] & ''))"
    $auditCount = 0
    foreach ($sourceKind in @("TreeClass", "TreeStatus")) {
        $source = Get-TreeHistoryConversionSourceDefinition -Connection $Connection -SourceKind $sourceKind
        if ($null -eq $source) { continue }

        foreach ($mode in @("Mismatch", "Review")) {
            $condition = Get-TreeHistoryConversionCondition `
                -SourceKind $sourceKind `
                -Mode $mode `
                -SourceExpression ([string]$source.SourceExpression) `
                -HistoryExpression $historyExpression
            if ([string]::IsNullOrWhiteSpace($condition)) { continue }

            $ruleName = Get-TreeHistoryConversionRuleName -SourceKind $sourceKind -Mode $mode
            $fieldName = Get-TreeHistoryConversionFieldName -SourceKind $sourceKind
            $message = Get-TreeHistoryConversionMessage `
                -SourceKind $sourceKind `
                -Mode $mode `
                -SourceTable ([string]$source.SourceTable) `
                -SourceField ([string]$source.SourceField)
            $tempName = "InventoryCleanXwalk_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $sourceLabel = Sql-Text "$($source.SourceTable).$($source.SourceField)"
            $selectSql = @"
SELECT
    tm.[MeasurementID] AS [RowID],
    $($source.RecordLabelExpression) AS [RecordLabel],
    ('$sourceKind=' & ($($source.SourceExpression) & '') & '; TreeHistory=' & ($historyExpression & '') & '; Source=' & $sourceLabel) AS [ObservedValue]
INTO $(Quote-Name $tempName)
$($source.FromSql)
WHERE $condition
"@
            $auditCount += Add-TempAuditFromSelect `
                -Connection $Connection `
                -TargetTableName "TreeMeasurements" `
                -TargetIdFieldName "MeasurementID" `
                -TempTableName $tempName `
                -SelectIntoSql $selectSql `
                -RuleName $ruleName `
                -FieldName $fieldName `
                -Message $message
        }
    }

    return $auditCount
}

function Get-TreeHistoryConversionVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$SourceKind,
        [string]$Mode
    )

    $source = Get-TreeHistoryConversionSourceDefinition -Connection $Connection -SourceKind $SourceKind
    if ($null -eq $source) { return "" }

    $historyExpression = "Trim((tm.[TreeHistory] & ''))"
    $condition = Get-TreeHistoryConversionCondition `
        -SourceKind $SourceKind `
        -Mode $Mode `
        -SourceExpression ([string]$source.SourceExpression) `
        -HistoryExpression $historyExpression
    if ([string]::IsNullOrWhiteSpace($condition)) { return "" }

    $tmSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "tm"
    $sourceFieldName = [string]$source.SourceField
    $extraColumns = ", $($source.SourceExpression) AS [Recorded${SourceKind}], $historyExpression AS [RecordedTreeHistory]"
    if ([string]$source.SourceTable -ne "TreeMeasurements") {
        $extraColumns += ", " + (Sql-Text "$($source.SourceTable).$sourceFieldName") + " AS [CrosswalkSourceField]"
    }

    return @"
SELECT $tmSelect$extraColumns
$($source.FromSql)
WHERE $condition;
"@
}

function Get-IngrowthPlotStatusSourceDefinition {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    foreach ($tableName in @("TreeMeasurements", "Trees", "PlotMeasurements")) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $tableName)) { return $null }
    }

    $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredField in @("MeasurementID", "TreeID", "PeriodNumber", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name $requiredField)) { return $null }
    }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "TreeHistory")) { return $null }

    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    foreach ($requiredField in @("TreeID", "PlotID")) {
        if (-not (Test-ColumnExists -Columns $treeColumns -Name $requiredField)) { return $null }
    }

    $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
    foreach ($requiredField in @("PlotID", "PeriodNumber")) {
        if (-not (Test-ColumnExists -Columns $plotMeasurementColumns -Name $requiredField)) { return $null }
    }

    $plotColumns = @()
    $canJoinPlots = $false
    if (Test-TableAvailable -Tables $tables -TableName "Plots") {
        try {
            $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
            $canJoinPlots = (Test-ColumnExists -Columns $plotColumns -Name "PlotID")
        }
        catch {
            $plotColumns = @()
            $canJoinPlots = $false
        }
    }

    $plotStatusExpression = Get-RequiredPlotStatusExpression `
        -MeasurementAlias "plotMeas" `
        -MeasurementColumns $plotMeasurementColumns `
        -PlotAlias $(if ($canJoinPlots) { "plotMeta" } else { "" }) `
        -PlotColumns $plotColumns
    if ([string]::IsNullOrWhiteSpace($plotStatusExpression)) { return $null }

    $fromSql = if ($canJoinPlots) {
        @"
FROM ((([TreeMeasurements] AS tm
INNER JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID])
INNER JOIN [PlotMeasurements] AS plotMeas
    ON treeMeta.[PlotID] = plotMeas.[PlotID])
INNER JOIN (
    SELECT [TreeID], Min([PeriodNumber]) AS [FirstTreePeriod]
    FROM [TreeMeasurements]
    GROUP BY [TreeID]
) AS firstTree
    ON tm.[TreeID] = firstTree.[TreeID])
LEFT JOIN [Plots] AS plotMeta
    ON treeMeta.[PlotID] = plotMeta.[PlotID]
"@
    }
    else {
        @"
FROM (([TreeMeasurements] AS tm
INNER JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID])
INNER JOIN [PlotMeasurements] AS plotMeas
    ON treeMeta.[PlotID] = plotMeas.[PlotID]
) INNER JOIN (
    SELECT [TreeID], Min([PeriodNumber]) AS [FirstTreePeriod]
    FROM [TreeMeasurements]
    GROUP BY [TreeID]
) AS firstTree
    ON tm.[TreeID] = firstTree.[TreeID]
"@
    }

    return [pscustomobject]@{
        FromSql = $fromSql
        PlotStatusExpression = $plotStatusExpression
        TreeHistoryExpression = "Trim((tm.[TreeHistory] & ''))"
        PriorExistsCondition = "EXISTS (SELECT * FROM [TreeMeasurements] AS priorTm WHERE priorTm.[TreeID] = tm.[TreeID] AND priorTm.[PeriodNumber] < tm.[PeriodNumber] AND Trim((priorTm.[TreeHistory] & '')) <> '')"
        FirstMeasurementCondition = "tm.[PeriodNumber] = firstTree.[FirstTreePeriod]"
        AfterInitialPeriodCondition = "tm.[PeriodNumber] > (SELECT Min(projectPeriod.[PeriodNumber]) FROM [TreeMeasurements] AS projectPeriod)"
        InitialProjectPeriodCondition = "tm.[PeriodNumber] = (SELECT Min(projectPeriod.[PeriodNumber]) FROM [TreeMeasurements] AS projectPeriod)"
        PeriodMatchCondition = "tm.[PeriodNumber] = plotMeas.[PeriodNumber]"
        RecordLabelExpression = (Get-TreeStemCountRecordLabelExpression -Columns $treeMeasurementColumns -Alias "tm")
        CanJoinPlots = [bool]$canJoinPlots
        UsesPlotMetaForPlotStatus = ([string]$plotStatusExpression -match "plotMeta\.")
    }
}

function Get-IngrowthCondition {
    param(
        [string]$RuleName,
        [string]$PlotStatusExpression,
        [string]$TreeHistoryExpression,
        [string]$PriorExistsCondition,
        [string]$FirstMeasurementCondition,
        [string]$AfterInitialPeriodCondition,
        [string]$InitialProjectPeriodCondition
    )

    switch ($RuleName) {
        "Ingrowth on install plot" {
            return "(Val($PlotStatusExpression) = 2 AND $TreeHistoryExpression = '10')"
        }
        "TreeHistory 10 in initial project period" {
            return "(Val($PlotStatusExpression) <> 2 AND $TreeHistoryExpression = '10' AND $InitialProjectPeriodCondition)"
        }
        "Ingrowth with prior tree measurement" {
            return "(Val($PlotStatusExpression) = 1 AND $TreeHistoryExpression = '10' AND $PriorExistsCondition)"
        }
        "Possible missed ingrowth" {
            return "(Val($PlotStatusExpression) = 1 AND $TreeHistoryExpression = '0' AND $FirstMeasurementCondition AND $AfterInitialPeriodCondition)"
        }
    }

    return ""
}

function Get-IngrowthRuleMessage {
    param([string]$RuleName)

    switch ($RuleName) {
        "Ingrowth on install plot" {
            return "TreeHistory 10 (ingrowth) should only be used for new trees on remeasurement plots. A new tree on an install plot with PlotStatus 2 should normally be TreeHistory 0."
        }
        "TreeHistory 10 in initial project period" {
            return "TreeHistory 10 (ingrowth) should not be used in the first project measurement period. First-period trees should normally be TreeHistory 0, then later new trees on remeasurement plots can be TreeHistory 10."
        }
        "Ingrowth with prior tree measurement" {
            return "TreeHistory 10 (ingrowth) is intended for the first nonblank TreeHistory record of a new tree on a remeasurement plot. This tree already has an earlier nonblank TreeHistory value, so review whether TreeHistory should be 0, 5, or another status."
        }
        "Possible missed ingrowth" {
            return "This appears to be the first recorded live TreeHistory 0 row for a tree on a remeasurement plot with PlotStatus 1. Review whether this tree should be TreeHistory 10 (ingrowth). Period 1 and install-plot trees should remain TreeHistory 0."
        }
    }

    return ""
}

function Get-PossibleMissedIngrowthPlotExistsFromSql {
    param([object]$Source)

    if ($null -ne $Source -and $Source.PSObject.Properties["UsesPlotMetaForPlotStatus"] -and [bool]$Source.UsesPlotMetaForPlotStatus) {
        return @"
FROM ([Trees] AS treeMeta
INNER JOIN [PlotMeasurements] AS plotMeas
    ON treeMeta.[PlotID] = plotMeas.[PlotID])
LEFT JOIN [Plots] AS plotMeta
    ON treeMeta.[PlotID] = plotMeta.[PlotID]
"@
    }

    return @"
FROM [Trees] AS treeMeta
INNER JOIN [PlotMeasurements] AS plotMeas
    ON treeMeta.[PlotID] = plotMeas.[PlotID]
"@
}

function Get-PossibleMissedIngrowthWhereSql {
    param(
        [object]$Source,
        [string]$FirstTreePeriodExpression
    )

    $existsFrom = Get-PossibleMissedIngrowthPlotExistsFromSql -Source $Source
    return @"
(
    $($Source.TreeHistoryExpression) = '0'
    AND tm.[PeriodNumber] = $FirstTreePeriodExpression
    AND $($Source.AfterInitialPeriodCondition)
    AND EXISTS (
        SELECT *
        $existsFrom
        WHERE treeMeta.[TreeID] = tm.[TreeID]
            AND plotMeas.[PeriodNumber] = tm.[PeriodNumber]
            AND Val($($Source.PlotStatusExpression)) = 1
    )
)
"@
}

function Add-PossibleMissedIngrowthCheck {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object]$Source
    )

    $message = Get-IngrowthRuleMessage -RuleName "Possible missed ingrowth"
    $firstTempName = "InventoryCleanFirstTree_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $firstTemp = Quote-Name $firstTempName

    try {
        $firstSql = @"
SELECT [TreeID], Min([PeriodNumber]) AS [FirstTreePeriod]
INTO $firstTemp
FROM [TreeMeasurements]
GROUP BY [TreeID]
"@
        [void](Invoke-NonQuery -Connection $Connection -Sql $firstSql)
        try { [void](Invoke-NonQuery -Connection $Connection -Sql "CREATE INDEX [idx_TreeID] ON $firstTemp ([TreeID])") } catch { }

        $whereCondition = Get-PossibleMissedIngrowthWhereSql -Source $Source -FirstTreePeriodExpression "firstTree.[FirstTreePeriod]"
        $tempName = "InventoryCleanIngrowth_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
        $selectSql = @"
SELECT
    tm.[MeasurementID] AS [RowID],
    $($Source.RecordLabelExpression) AS [RecordLabel],
    ('PlotStatus=1; TreeHistory=' & ($($Source.TreeHistoryExpression) & '') & '; Possible missed ingrowth') AS [ObservedValue]
INTO $(Quote-Name $tempName)
FROM [TreeMeasurements] AS tm
INNER JOIN $firstTemp AS firstTree
    ON tm.[TreeID] = firstTree.[TreeID]
WHERE $whereCondition
"@
        return Add-TempAuditFromSelect `
            -Connection $Connection `
            -TargetTableName "TreeMeasurements" `
            -TargetIdFieldName "MeasurementID" `
            -TempTableName $tempName `
            -SelectIntoSql $selectSql `
            -RuleName "Possible missed ingrowth" `
            -FieldName "TreeHistory" `
            -Message $message
    }
    finally {
        try { [void](Invoke-NonQuery -Connection $Connection -Sql "DROP TABLE $firstTemp") } catch { }
    }
}

function Add-IngrowthTreeHistoryChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $source = Get-IngrowthPlotStatusSourceDefinition -Connection $Connection
    if ($null -eq $source) { return 0 }

    Ensure-NeedsReviewColumn -Connection $Connection -TableName "TreeMeasurements"
    $auditCount = 0
    foreach ($ruleName in @("Ingrowth on install plot", "TreeHistory 10 in initial project period", "Ingrowth with prior tree measurement", "Possible missed ingrowth")) {
        if ($ruleName -eq "Possible missed ingrowth") {
            $auditCount += Add-PossibleMissedIngrowthCheck -Connection $Connection -Source $source
            continue
        }

        $condition = Get-IngrowthCondition `
            -RuleName $ruleName `
            -PlotStatusExpression ([string]$source.PlotStatusExpression) `
            -TreeHistoryExpression ([string]$source.TreeHistoryExpression) `
            -PriorExistsCondition ([string]$source.PriorExistsCondition) `
            -FirstMeasurementCondition ([string]$source.FirstMeasurementCondition) `
            -AfterInitialPeriodCondition ([string]$source.AfterInitialPeriodCondition) `
            -InitialProjectPeriodCondition ([string]$source.InitialProjectPeriodCondition)
        if ([string]::IsNullOrWhiteSpace($condition)) { continue }
        $whereCondition = "($($source.PeriodMatchCondition) AND $condition)"

        $message = Get-IngrowthRuleMessage -RuleName $ruleName
        $tempName = "InventoryCleanIngrowth_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
        $selectSql = @"
SELECT
    tm.[MeasurementID] AS [RowID],
    $($source.RecordLabelExpression) AS [RecordLabel],
    ('PlotStatus=' & ($($source.PlotStatusExpression) & '') & '; TreeHistory=' & ($($source.TreeHistoryExpression) & '') & '; $ruleName') AS [ObservedValue]
INTO $(Quote-Name $tempName)
$($source.FromSql)
WHERE $whereCondition
"@
        $auditCount += Add-TempAuditFromSelect `
            -Connection $Connection `
            -TargetTableName "TreeMeasurements" `
            -TargetIdFieldName "MeasurementID" `
            -TempTableName $tempName `
            -SelectIntoSql $selectSql `
            -RuleName $ruleName `
            -FieldName "TreeHistory" `
            -Message $message
    }

    return $auditCount
}

function Get-IngrowthVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$RuleName
    )

    $source = Get-IngrowthPlotStatusSourceDefinition -Connection $Connection
    if ($null -eq $source) { return "" }

    if ($RuleName -eq "Possible missed ingrowth") {
        $tmSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "tm"
        $whereCondition = Get-PossibleMissedIngrowthWhereSql `
            -Source $source `
            -FirstTreePeriodExpression "(SELECT Min(ft.[PeriodNumber]) FROM [TreeMeasurements] AS ft WHERE ft.[TreeID] = tm.[TreeID])"
        return @"
SELECT $tmSelect, 1 AS [ResolvedPlotStatus]
FROM [TreeMeasurements] AS tm
WHERE $whereCondition;
"@
    }

    $condition = Get-IngrowthCondition `
        -RuleName $RuleName `
        -PlotStatusExpression ([string]$source.PlotStatusExpression) `
        -TreeHistoryExpression ([string]$source.TreeHistoryExpression) `
        -PriorExistsCondition ([string]$source.PriorExistsCondition) `
        -FirstMeasurementCondition ([string]$source.FirstMeasurementCondition) `
        -AfterInitialPeriodCondition ([string]$source.AfterInitialPeriodCondition) `
        -InitialProjectPeriodCondition ([string]$source.InitialProjectPeriodCondition)
    if ([string]::IsNullOrWhiteSpace($condition)) { return "" }
    $whereCondition = "($($source.PeriodMatchCondition) AND $condition)"

    $tmSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "tm"
    return @"
SELECT $tmSelect, $($source.PlotStatusExpression) AS [ResolvedPlotStatus]
$($source.FromSql)
WHERE $whereCondition;
"@
}

function Test-NoTreeDataSystemField {
    param([string]$FieldName)

    $field = Normalize-FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($field)) { return $true }
    if ($field -in @(
        "measurementid",
        "treeid",
        "treekey",
        "treemeaskey",
        "periodnumber",
        "projectperiodid",
        "plotid",
        "plotkey",
        "plotnumber",
        "treenumber",
        "needsreview",
        "auditid",
        "audittime",
        "sourcerowid",
        "rowid",
        "createddate",
        "createdby",
        "modifieddate",
        "modifiedby",
        "updateddate",
        "updatedby",
        "lastmodifieddate",
        "lastmodifiedby"
    )) { return $true }
    if ($field -match "remark|remarks|note|notes|comment|comments") { return $true }
    return $false
}

function Get-NoTreeDataReviewFields {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [object[]]$Columns
    )

    $fields = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $metadata = Get-AppColumnReviewMetadata -Connection $Connection
    if ([bool]$metadata.HasMetadata) {
        foreach ($item in @($metadata.Rows)) {
            if (-not [bool]$item.IsActive) { continue }
            if (-not ([string]$item.TableName).Equals($TableName, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $fieldName = [string]$item.FieldName
            if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
            if (Test-NoTreeDataSystemField -FieldName $fieldName) { continue }
            if (Test-ExcludedCleaningField -FieldName $fieldName) { continue }
            if (-not (Test-ColumnExists -Columns $Columns -Name $fieldName)) { continue }
            if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName $fieldName)) { continue }

            $key = Normalize-FieldName $fieldName
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$fields.Add($fieldName)
            }
        }
    }
    elseif ($TableName.Equals("TreeMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) {
        foreach ($column in $Columns) {
            $fieldName = [string]$column.Name
            if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
            if (Test-NoTreeDataSystemField -FieldName $fieldName) { continue }
            if (-not (Test-SupportedDataEntryReviewField -TableName $TableName -FieldName $fieldName)) { continue }

            $key = Normalize-FieldName $fieldName
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$fields.Add($fieldName)
            }
        }
    }

    return [string[]]$fields.ToArray()
}

function Get-NoTreeDataRecordedCondition {
    param(
        [string]$FieldExpression,
        [string]$FieldName
    )

    $valueText = "Trim(($FieldExpression & ''))"
    $field = Normalize-FieldName $FieldName
    if ($field -in @("problem1", "problem2", "severity1", "severity2")) {
        return "($FieldExpression Is Not Null AND $valueText <> '' AND Val($valueText) > 0)"
    }

    return "($FieldExpression Is Not Null AND $valueText <> '')"
}

function Get-NoTreeDataOnInactivePlotStatusMessage {
    return "Tree data should not be recorded when the matching plot status is 4 (missing plot), 5 (not measured), 6 (dropped - off reservation), or 7 (dropped - other). Clear the recorded tree value, or correct PlotStatus if the plot was actually measured."
}

function Get-NoTreeDataOnInactivePlotStatusSourceDefinition {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    foreach ($requiredTable in @("TreeMeasurements", "Trees", "PlotMeasurements")) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $requiredTable)) { return $null }
    }
    if (-not (Test-TableAvailable -Tables $tables -TableName $TableName)) { return $null }

    $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredField in @("TreeID", "PeriodNumber")) {
        if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name $requiredField)) { return $null }
    }
    if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name "TreeMeasKey") -and
        $TableName.Equals("TreeCustomMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) { return $null }

    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    foreach ($requiredField in @("TreeID", "PlotID")) {
        if (-not (Test-ColumnExists -Columns $treeColumns -Name $requiredField)) { return $null }
    }

    $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
    foreach ($requiredField in @("PlotID", "PeriodNumber")) {
        if (-not (Test-ColumnExists -Columns $plotMeasurementColumns -Name $requiredField)) { return $null }
    }

    $plotColumns = @()
    $canJoinPlots = $false
    if (Test-TableAvailable -Tables $tables -TableName "Plots") {
        try {
            $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
            $canJoinPlots = (Test-ColumnExists -Columns $plotColumns -Name "PlotID")
        }
        catch {
            $plotColumns = @()
            $canJoinPlots = $false
        }
    }

    $plotStatusExpression = Get-RequiredPlotStatusExpression `
        -MeasurementAlias "plotMeas" `
        -MeasurementColumns $plotMeasurementColumns `
        -PlotAlias $(if ($canJoinPlots) { "plotMeta" } else { "" }) `
        -PlotColumns $plotColumns
    if ([string]::IsNullOrWhiteSpace($plotStatusExpression)) { return $null }

    $sourceColumns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    $sourceAlias = "tm"
    $sourceIdField = "MeasurementID"
    $targetIdField = "MeasurementID"
    $sourceIdExpression = "tm.[MeasurementID]"
    if ($TableName.Equals("TreeCustomMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) {
        $sourceAlias = "c"
        foreach ($requiredField in @("MeasurementID", "TreeMeasKey")) {
            if (-not (Test-ColumnExists -Columns $sourceColumns -Name $requiredField)) { return $null }
        }
        $sourceIdExpression = "c.[MeasurementID]"
    }
    elseif (-not (Test-ColumnExists -Columns $sourceColumns -Name "MeasurementID")) {
        return $null
    }

    $fields = @(Get-NoTreeDataReviewFields -Connection $Connection -TableName $TableName -Columns $sourceColumns)
    if ($fields.Count -eq 0) { return $null }

    $fromSql = ""
    if ($TableName.Equals("TreeCustomMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) {
        $fromSql = if ($canJoinPlots) {
            @"
FROM ((([TreeCustomMeasurements] AS c
INNER JOIN [TreeMeasurements] AS tm
    ON c.[TreeMeasKey] = tm.[TreeMeasKey])
INNER JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID])
INNER JOIN [PlotMeasurements] AS plotMeas
    ON treeMeta.[PlotID] = plotMeas.[PlotID])
LEFT JOIN [Plots] AS plotMeta
    ON treeMeta.[PlotID] = plotMeta.[PlotID]
"@
        }
        else {
            @"
FROM (([TreeCustomMeasurements] AS c
INNER JOIN [TreeMeasurements] AS tm
    ON c.[TreeMeasKey] = tm.[TreeMeasKey])
INNER JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID])
INNER JOIN [PlotMeasurements] AS plotMeas
    ON treeMeta.[PlotID] = plotMeas.[PlotID]
"@
        }
    }
    else {
        $fromSql = if ($canJoinPlots) {
            @"
FROM (([TreeMeasurements] AS tm
INNER JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID])
INNER JOIN [PlotMeasurements] AS plotMeas
    ON treeMeta.[PlotID] = plotMeas.[PlotID])
LEFT JOIN [Plots] AS plotMeta
    ON treeMeta.[PlotID] = plotMeta.[PlotID]
"@
        }
        else {
            @"
FROM ([TreeMeasurements] AS tm
INNER JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID])
INNER JOIN [PlotMeasurements] AS plotMeas
    ON treeMeta.[PlotID] = plotMeas.[PlotID]
"@
        }
    }

    $treeHistoryObserved = if (Test-ColumnExists -Columns $treeMeasurementColumns -Name "TreeHistory") { "(tm.[TreeHistory] & '')" } else { "''" }

    return [pscustomobject]@{
        TableName = $TableName
        Fields = $fields
        FromSql = $fromSql
        SourceAlias = $sourceAlias
        SourceIdField = $sourceIdField
        TargetIdField = $targetIdField
        SourceIdExpression = $sourceIdExpression
        PlotStatusExpression = $plotStatusExpression
        PeriodMatchCondition = "tm.[PeriodNumber] = plotMeas.[PeriodNumber]"
        InactivePlotStatusCondition = "Val($plotStatusExpression) In (4, 5, 6, 7)"
        PeriodScopeCondition = (Get-SelectedPeriodAliasCondition -Alias "tm")
        RecordLabelExpression = (Get-TreeStemCountRecordLabelExpression -Columns $treeMeasurementColumns -Alias "tm")
        TreeHistoryObserved = $treeHistoryObserved
    }
}

function Add-NoTreeDataOnInactivePlotStatusChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $auditCount = 0
    foreach ($tableName in @("TreeMeasurements", "TreeCustomMeasurements")) {
        $source = Get-NoTreeDataOnInactivePlotStatusSourceDefinition -Connection $Connection -TableName $tableName
        if ($null -eq $source) { continue }

        Ensure-NeedsReviewColumn -Connection $Connection -TableName $tableName
        foreach ($fieldName in @($source.Fields)) {
            $fieldExpression = "$($source.SourceAlias).$(Quote-Name $fieldName)"
            $recordedCondition = Get-NoTreeDataRecordedCondition -FieldExpression $fieldExpression -FieldName $fieldName
            $whereParts = New-Object System.Collections.Generic.List[string]
            [void]$whereParts.Add([string]$source.PeriodMatchCondition)
            [void]$whereParts.Add([string]$source.InactivePlotStatusCondition)
            [void]$whereParts.Add($recordedCondition)
            if (-not [string]::IsNullOrWhiteSpace([string]$source.PeriodScopeCondition)) {
                [void]$whereParts.Add([string]$source.PeriodScopeCondition)
            }
            $whereCondition = "(" + ([string]::Join(" AND ", [string[]]$whereParts.ToArray())) + ")"

            $tempName = "InventoryCleanNoTreeData_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $selectSql = @"
SELECT
    $($source.SourceIdExpression) AS [RowID],
    $($source.RecordLabelExpression) AS [RecordLabel],
    ('PlotStatus=' & ($($source.PlotStatusExpression) & '') & '; $fieldName=' & ($fieldExpression & '') & '; TreeHistory=' & ($($source.TreeHistoryObserved) & '')) AS [ObservedValue]
INTO $(Quote-Name $tempName)
$($source.FromSql)
WHERE $whereCondition
"@

            $auditCount += Add-TempAuditFromSelect `
                -Connection $Connection `
                -TargetTableName $tableName `
                -TargetIdFieldName ([string]$source.TargetIdField) `
                -TempTableName $tempName `
                -SelectIntoSql $selectSql `
                -RuleName "Tree data on not-measured plot" `
                -FieldName $fieldName `
                -Message (Get-NoTreeDataOnInactivePlotStatusMessage)
        }
    }

    return $auditCount
}

function Get-NoTreeDataOnInactivePlotStatusVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName
    )

    $source = Get-NoTreeDataOnInactivePlotStatusSourceDefinition -Connection $Connection -TableName $TableName
    if ($null -eq $source) { return "" }
    if (-not (@($source.Fields) -contains $FieldName)) { return "" }

    $fieldExpression = "$($source.SourceAlias).$(Quote-Name $FieldName)"
    $recordedCondition = Get-NoTreeDataRecordedCondition -FieldExpression $fieldExpression -FieldName $FieldName
    $whereParts = New-Object System.Collections.Generic.List[string]
    [void]$whereParts.Add([string]$source.PeriodMatchCondition)
    [void]$whereParts.Add([string]$source.InactivePlotStatusCondition)
    [void]$whereParts.Add($recordedCondition)
    if (-not [string]::IsNullOrWhiteSpace([string]$source.PeriodScopeCondition)) {
        [void]$whereParts.Add([string]$source.PeriodScopeCondition)
    }
    $whereCondition = [string]::Join(" AND ", [string[]]$whereParts.ToArray())

    $selectList = Get-AliasSelectList -Connection $Connection -TableName $TableName -Alias ([string]$source.SourceAlias)
    return @"
SELECT $selectList, $($source.PlotStatusExpression) AS [ResolvedPlotStatus], tm.[PeriodNumber] AS [TreeMeasurementPeriod]
$($source.FromSql)
WHERE $whereCondition;
"@
}

function Test-NoPlotDataAllowedOrSystemField {
    param([string]$FieldName)

    $field = Normalize-FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($field)) { return $true }
    if ($field -in @(
        "measurementid",
        "plotid",
        "plotkey",
        "plotmeaskey",
        "periodnumber",
        "projectperiodid",
        "needsreview",
        "auditid",
        "audittime",
        "sourcerowid",
        "rowid",
        "created",
        "createddate",
        "createdby",
        "updated",
        "updateddate",
        "updatedby",
        "modifieddate",
        "modifiedby",
        "lastmodifieddate",
        "lastmodifiedby",
        "retrievalid",
        "plotstatus",
        "plotstatuscode",
        "plotstatusid",
        "status",
        "statuscode",
        "statusid"
    )) { return $true }

    if ($field -match "remark|remarks|note|notes|comment|comments") { return $true }
    if ($field -in @("utmnorthing", "northing", "utmeasting", "easting", "utmzone", "utmzonecode", "zone")) { return $true }
    if ($field -match "^utm") { return $true }
    if ($field -in @("managementunit", "managementunitcode", "managementunitid", "managmentunit", "managmentunitcode", "mgmtunit", "mgmtunitcode", "mgtunit", "mgtunitcode", "mu", "mucode")) { return $true }
    if ($field -eq "flccommercial" -or ($field -match "flc" -and $field -match "commercial")) { return $true }

    return $false
}

function Get-NoPlotDataReviewFields {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [object[]]$Columns
    )

    $fields = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $metadata = Get-AppColumnReviewMetadata -Connection $Connection
    if ([bool]$metadata.HasMetadata) {
        foreach ($item in @($metadata.Rows)) {
            if (-not [bool]$item.IsActive) { continue }
            if (-not ([string]$item.TableName).Equals($TableName, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $fieldName = [string]$item.FieldName
            if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
            if (Test-NoPlotDataAllowedOrSystemField -FieldName $fieldName) { continue }
            if (Test-ExcludedCleaningField -FieldName $fieldName) { continue }
            if (-not (Test-ColumnExists -Columns $Columns -Name $fieldName)) { continue }
            if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName $fieldName)) { continue }

            $key = Normalize-FieldName $fieldName
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$fields.Add($fieldName)
            }
        }
    }
    elseif ($TableName.Equals("PlotMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) {
        foreach ($column in $Columns) {
            $fieldName = [string]$column.Name
            if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
            if (Test-NoPlotDataAllowedOrSystemField -FieldName $fieldName) { continue }
            if (Test-ExcludedCleaningField -FieldName $fieldName) { continue }
            if (-not (Test-SupportedDataEntryReviewField -TableName $TableName -FieldName $fieldName)) { continue }

            $key = Normalize-FieldName $fieldName
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$fields.Add($fieldName)
            }
        }
    }

    return [string[]]$fields.ToArray()
}

function Get-NoPlotDataRecordedCondition {
    param([string]$FieldExpression)

    $valueText = "Trim(($FieldExpression & ''))"
    return "($FieldExpression Is Not Null AND $valueText <> '')"
}

function Get-NoPlotDataOnInactivePlotStatusMessage {
    return "Plot data should not be recorded when PlotStatus is 4 (missing plot), 5 (not measured), 6 (dropped - off reservation), or 7 (dropped - other). Plot remarks, UTM coordinates, management unit, and FLCCommercial are allowed because they may be preloaded or needed to explain the dropped plot."
}

function Get-NoPlotDataOnInactivePlotStatusSourceDefinition {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    foreach ($requiredTable in @("PlotMeasurements", $TableName)) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $requiredTable)) { return $null }
    }

    $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
    foreach ($requiredField in @("PlotID", "PeriodNumber")) {
        if (-not (Test-ColumnExists -Columns $plotMeasurementColumns -Name $requiredField)) { return $null }
    }

    $plotColumns = @()
    $canJoinPlots = $false
    if (Test-TableAvailable -Tables $tables -TableName "Plots") {
        try {
            $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
            $canJoinPlots = (Test-ColumnExists -Columns $plotColumns -Name "PlotID")
        }
        catch {
            $plotColumns = @()
            $canJoinPlots = $false
        }
    }

    $plotStatusExpression = Get-RequiredPlotStatusExpression `
        -MeasurementAlias "pm" `
        -MeasurementColumns $plotMeasurementColumns `
        -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
        -PlotColumns $plotColumns
    if ([string]::IsNullOrWhiteSpace($plotStatusExpression)) { return $null }

    $sourceColumns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    $sourceAlias = "pm"
    $targetIdField = Get-FirstMatchingColumnName -Columns $sourceColumns -CandidateNames @("MeasurementID", "ID")
    if ([string]::IsNullOrWhiteSpace($targetIdField)) { return $null }

    if ($TableName.Equals("PlotCustomMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) {
        $sourceAlias = "c"
        foreach ($requiredField in @("PlotMeasKey")) {
            if (-not (Test-ColumnExists -Columns $sourceColumns -Name $requiredField)) { return $null }
            if (-not (Test-ColumnExists -Columns $plotMeasurementColumns -Name $requiredField)) { return $null }
        }
    }

    $fields = @(Get-NoPlotDataReviewFields -Connection $Connection -TableName $TableName -Columns $sourceColumns)
    if ($fields.Count -eq 0) { return $null }

    $fromSql = ""
    if ($TableName.Equals("PlotCustomMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) {
        $fromSql = if ($canJoinPlots) {
            @"
FROM ([PlotCustomMeasurements] AS c
INNER JOIN [PlotMeasurements] AS pm
    ON c.[PlotMeasKey] = pm.[PlotMeasKey])
LEFT JOIN [Plots] AS pl
    ON pm.[PlotID] = pl.[PlotID]
"@
        }
        else {
            @"
FROM [PlotCustomMeasurements] AS c
INNER JOIN [PlotMeasurements] AS pm
    ON c.[PlotMeasKey] = pm.[PlotMeasKey]
"@
        }
    }
    else {
        $fromSql = if ($canJoinPlots) {
            @"
FROM [PlotMeasurements] AS pm
LEFT JOIN [Plots] AS pl
    ON pm.[PlotID] = pl.[PlotID]
"@
        }
        else {
            "FROM [PlotMeasurements] AS pm"
        }
    }

    return [pscustomobject]@{
        TableName = $TableName
        Fields = $fields
        FromSql = $fromSql
        SourceAlias = $sourceAlias
        SourceIdField = $targetIdField
        TargetIdField = $targetIdField
        SourceIdExpression = "$sourceAlias.$(Quote-Name $targetIdField)"
        PlotStatusExpression = $plotStatusExpression
        InactivePlotStatusCondition = "Val($plotStatusExpression) In (4, 5, 6, 7)"
        PeriodScopeCondition = (Get-SelectedPeriodAliasCondition -Alias "pm")
        RecordLabelExpression = (Get-RequiredPlotRecordLabelExpression `
            -MeasurementAlias "pm" `
            -MeasurementColumns $plotMeasurementColumns `
            -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
            -PlotColumns $plotColumns `
            -FallbackExpression "($sourceAlias.$(Quote-Name $targetIdField) & '')")
    }
}

function Add-NoPlotDataOnInactivePlotStatusChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $auditCount = 0
    foreach ($tableName in @("PlotMeasurements", "PlotCustomMeasurements")) {
        $source = Get-NoPlotDataOnInactivePlotStatusSourceDefinition -Connection $Connection -TableName $tableName
        if ($null -eq $source) { continue }

        Ensure-NeedsReviewColumn -Connection $Connection -TableName $tableName
        foreach ($fieldName in @($source.Fields)) {
            $fieldExpression = "$($source.SourceAlias).$(Quote-Name $fieldName)"
            $recordedCondition = Get-NoPlotDataRecordedCondition -FieldExpression $fieldExpression
            $whereParts = New-Object System.Collections.Generic.List[string]
            [void]$whereParts.Add([string]$source.InactivePlotStatusCondition)
            [void]$whereParts.Add($recordedCondition)
            if (-not [string]::IsNullOrWhiteSpace([string]$source.PeriodScopeCondition)) {
                [void]$whereParts.Add([string]$source.PeriodScopeCondition)
            }
            $whereCondition = "(" + ([string]::Join(" AND ", [string[]]$whereParts.ToArray())) + ")"

            $tempName = "InventoryCleanNoPlotData_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $selectSql = @"
SELECT
    $($source.SourceIdExpression) AS [RowID],
    $($source.RecordLabelExpression) AS [RecordLabel],
    ('PlotStatus=' & ($($source.PlotStatusExpression) & '') & '; $fieldName=' & ($fieldExpression & '')) AS [ObservedValue]
INTO $(Quote-Name $tempName)
$($source.FromSql)
WHERE $whereCondition
"@

            $auditCount += Add-TempAuditFromSelect `
                -Connection $Connection `
                -TargetTableName $tableName `
                -TargetIdFieldName ([string]$source.TargetIdField) `
                -TempTableName $tempName `
                -SelectIntoSql $selectSql `
                -RuleName "Plot data on not-measured plot" `
                -FieldName $fieldName `
                -Message (Get-NoPlotDataOnInactivePlotStatusMessage)
        }
    }

    return $auditCount
}

function Get-NoPlotDataOnInactivePlotStatusVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName
    )

    $source = Get-NoPlotDataOnInactivePlotStatusSourceDefinition -Connection $Connection -TableName $TableName
    if ($null -eq $source) { return "" }
    if (-not (@($source.Fields) -contains $FieldName)) { return "" }

    $fieldExpression = "$($source.SourceAlias).$(Quote-Name $FieldName)"
    $recordedCondition = Get-NoPlotDataRecordedCondition -FieldExpression $fieldExpression
    $whereParts = New-Object System.Collections.Generic.List[string]
    [void]$whereParts.Add([string]$source.InactivePlotStatusCondition)
    [void]$whereParts.Add($recordedCondition)
    if (-not [string]::IsNullOrWhiteSpace([string]$source.PeriodScopeCondition)) {
        [void]$whereParts.Add([string]$source.PeriodScopeCondition)
    }
    $whereCondition = [string]::Join(" AND ", [string[]]$whereParts.ToArray())

    $selectList = Get-AliasSelectList -Connection $Connection -TableName $TableName -Alias ([string]$source.SourceAlias)
    return @"
SELECT $selectList, $($source.PlotStatusExpression) AS [ResolvedPlotStatus], pm.[PeriodNumber] AS [PlotMeasurementPeriod]
$($source.FromSql)
WHERE $whereCondition;
"@
}

function Test-NoRegenDataAllowedOrSystemField {
    param([string]$FieldName)

    $field = Normalize-FieldName $FieldName
    if ([string]::IsNullOrWhiteSpace($field)) { return $true }
    if ($field -in @(
        "measurementid",
        "regenmeaskey",
        "plotid",
        "plotkey",
        "plotnumber",
        "projectperiodid",
        "periodnumber",
        "needsreview",
        "auditid",
        "audittime",
        "sourcerowid",
        "rowid",
        "created",
        "createddate",
        "createdby",
        "updated",
        "updateddate",
        "updatedby",
        "modifieddate",
        "modifiedby",
        "lastmodifieddate",
        "lastmodifiedby",
        "upduser",
        "retrievalid",
        "isdeleted",
        "expfactor",
        "expansionfactor"
    )) { return $true }

    if ($field -match "remark|remarks|note|notes|comment|comments") { return $true }
    return $false
}

function Get-NoRegenDataReviewFields {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [object[]]$Columns
    )

    $fields = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $metadata = Get-AppColumnReviewMetadata -Connection $Connection
    if ([bool]$metadata.HasMetadata) {
        foreach ($item in @($metadata.Rows)) {
            if (-not [bool]$item.IsActive) { continue }
            if (-not ([string]$item.TableName).Equals($TableName, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $fieldName = [string]$item.FieldName
            if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
            if (Test-NoRegenDataAllowedOrSystemField -FieldName $fieldName) { continue }
            if (Test-ExcludedCleaningField -FieldName $fieldName) { continue }
            if (-not (Test-ColumnExists -Columns $Columns -Name $fieldName)) { continue }
            if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName $fieldName)) { continue }

            $key = Normalize-FieldName $fieldName
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$fields.Add($fieldName)
            }
        }
    }
    else {
        foreach ($column in $Columns) {
            $fieldName = [string]$column.Name
            if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
            if (Test-NoRegenDataAllowedOrSystemField -FieldName $fieldName) { continue }
            if (Test-ExcludedCleaningField -FieldName $fieldName) { continue }
            if ($TableName.Equals("RegenMeasurements", [System.StringComparison]::OrdinalIgnoreCase) -and
                -not (Test-SupportedDataEntryReviewField -TableName $TableName -FieldName $fieldName)) { continue }

            $key = Normalize-FieldName $fieldName
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$fields.Add($fieldName)
            }
        }
    }

    return [string[]]$fields.ToArray()
}

function Get-NoRegenDataRecordedCondition {
    param([string]$FieldExpression)

    $valueText = "Trim(($FieldExpression & ''))"
    return "($FieldExpression Is Not Null AND $valueText <> '')"
}

function Get-NoRegenDataOnInactivePlotStatusMessage {
    return "Regen data should not be recorded when the matching plot status is 4 (missing plot), 5 (not measured), 6 (dropped - off reservation), or 7 (dropped - other). Clear the recorded regen value, or correct PlotStatus if the plot was actually measured."
}

function Get-NoRegenDataOnInactivePlotStatusSourceDefinition {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    foreach ($requiredTable in @("RegenMeasurements", "PlotMeasurements", $TableName)) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $requiredTable)) { return $null }
    }

    $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements")
    foreach ($requiredField in @("MeasurementID", "RegenMeasKey")) {
        if (-not (Test-ColumnExists -Columns $regenColumns -Name $requiredField)) { return $null }
    }

    $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
    if (-not (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PeriodNumber")) { return $null }
    $plotJoinField = Get-FirstMatchingColumnName -Columns $regenColumns -CandidateNames @("PlotID", "PlotKey", "PlotNumber")
    if ([string]::IsNullOrWhiteSpace($plotJoinField) -or
        -not (Test-ColumnExists -Columns $plotMeasurementColumns -Name $plotJoinField)) { return $null }

    $periodJoinCondition = ""
    if ((Test-ColumnExists -Columns $regenColumns -Name "ProjectPeriodID") -and
        (Test-ColumnExists -Columns $plotMeasurementColumns -Name "ProjectPeriodID")) {
        $periodJoinCondition = "r.[ProjectPeriodID] = pm.[ProjectPeriodID]"
    }
    elseif ((Test-ColumnExists -Columns $regenColumns -Name "PeriodNumber") -and
        (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PeriodNumber")) {
        $periodJoinCondition = "r.[PeriodNumber] = pm.[PeriodNumber]"
    }
    elseif (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PeriodNumber") {
        $periodJoinCondition = "$(Get-RegenMeasKeyPeriodValueExpression -Alias "r") = pm.[PeriodNumber]"
    }
    else {
        return $null
    }

    $plotColumns = @()
    $canJoinPlots = $false
    if (Test-TableAvailable -Tables $tables -TableName "Plots") {
        try {
            $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
            $canJoinPlots = (Test-ColumnExists -Columns $plotColumns -Name "PlotID") -and
                (Test-ColumnExists -Columns $regenColumns -Name "PlotID")
        }
        catch {
            $plotColumns = @()
            $canJoinPlots = $false
        }
    }

    $plotStatusExpression = Get-RequiredPlotStatusExpression `
        -MeasurementAlias "pm" `
        -MeasurementColumns $plotMeasurementColumns `
        -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
        -PlotColumns $plotColumns
    if ([string]::IsNullOrWhiteSpace($plotStatusExpression)) { return $null }

    $sourceColumns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    $sourceAlias = "r"
    $targetIdField = Get-FirstMatchingColumnName -Columns $sourceColumns -CandidateNames @("MeasurementID", "ID")
    if ([string]::IsNullOrWhiteSpace($targetIdField)) { return $null }

    if ($TableName.Equals("RegenCustomMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) {
        $sourceAlias = "c"
        if (-not (Test-ColumnExists -Columns $sourceColumns -Name "RegenMeasKey")) { return $null }
    }

    $fields = @(Get-NoRegenDataReviewFields -Connection $Connection -TableName $TableName -Columns $sourceColumns)
    if ($fields.Count -eq 0) { return $null }

    $periodMatchCondition = "r.$(Quote-Name $plotJoinField) = pm.$(Quote-Name $plotJoinField) AND $periodJoinCondition"
    $fromSql = ""
    if ($TableName.Equals("RegenCustomMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) {
        $fromSql = if ($canJoinPlots) {
            @"
FROM (([RegenCustomMeasurements] AS c
INNER JOIN [RegenMeasurements] AS r
    ON c.[RegenMeasKey] = r.[RegenMeasKey])
INNER JOIN [PlotMeasurements] AS pm
    ON $periodMatchCondition)
LEFT JOIN [Plots] AS pl
    ON r.[PlotID] = pl.[PlotID]
"@
        }
        else {
            @"
FROM ([RegenCustomMeasurements] AS c
INNER JOIN [RegenMeasurements] AS r
    ON c.[RegenMeasKey] = r.[RegenMeasKey])
INNER JOIN [PlotMeasurements] AS pm
    ON $periodMatchCondition
"@
        }
    }
    else {
        $fromSql = if ($canJoinPlots) {
            @"
FROM ([RegenMeasurements] AS r
INNER JOIN [PlotMeasurements] AS pm
    ON $periodMatchCondition)
LEFT JOIN [Plots] AS pl
    ON r.[PlotID] = pl.[PlotID]
"@
        }
        else {
            @"
FROM [RegenMeasurements] AS r
INNER JOIN [PlotMeasurements] AS pm
    ON $periodMatchCondition
"@
        }
    }

    $recordLabelExpression = (Get-RequiredPlotRecordLabelExpression `
        -MeasurementAlias "pm" `
        -MeasurementColumns $plotMeasurementColumns `
        -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
        -PlotColumns $plotColumns `
        -FallbackExpression "(r.[RegenMeasKey] & '')")
    $recordLabelExpression = "($recordLabelExpression & ' Regen ' & (r.[RegenMeasKey] & ''))"

    return [pscustomobject]@{
        TableName = $TableName
        Fields = $fields
        FromSql = $fromSql
        SourceAlias = $sourceAlias
        SourceIdField = $targetIdField
        TargetIdField = $targetIdField
        SourceIdExpression = "$sourceAlias.$(Quote-Name $targetIdField)"
        PlotStatusExpression = $plotStatusExpression
        PeriodMatchCondition = $periodMatchCondition
        InactivePlotStatusCondition = "Val($plotStatusExpression) In (4, 5, 6, 7)"
        PeriodScopeCondition = (Get-RegenMeasKeyPeriodScopeCondition -Alias "r")
        RecordLabelExpression = $recordLabelExpression
    }
}

function Add-NoRegenDataOnInactivePlotStatusChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $auditCount = 0
    foreach ($tableName in @("RegenMeasurements", "RegenCustomMeasurements")) {
        $source = Get-NoRegenDataOnInactivePlotStatusSourceDefinition -Connection $Connection -TableName $tableName
        if ($null -eq $source) { continue }

        Ensure-NeedsReviewColumn -Connection $Connection -TableName $tableName
        foreach ($fieldName in @($source.Fields)) {
            $fieldExpression = "$($source.SourceAlias).$(Quote-Name $fieldName)"
            $recordedCondition = Get-NoRegenDataRecordedCondition -FieldExpression $fieldExpression
            $whereParts = New-Object System.Collections.Generic.List[string]
            [void]$whereParts.Add([string]$source.InactivePlotStatusCondition)
            [void]$whereParts.Add($recordedCondition)
            if (-not [string]::IsNullOrWhiteSpace([string]$source.PeriodScopeCondition)) {
                [void]$whereParts.Add([string]$source.PeriodScopeCondition)
            }
            $whereCondition = "(" + ([string]::Join(" AND ", [string[]]$whereParts.ToArray())) + ")"

            $tempName = "InventoryCleanNoRegenData_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $selectSql = @"
SELECT
    $($source.SourceIdExpression) AS [RowID],
    $($source.RecordLabelExpression) AS [RecordLabel],
    ('PlotStatus=' & ($($source.PlotStatusExpression) & '') & '; $fieldName=' & ($fieldExpression & '') & '; RegenMeasKey=' & (r.[RegenMeasKey] & '')) AS [ObservedValue]
INTO $(Quote-Name $tempName)
$($source.FromSql)
WHERE $whereCondition
"@

            $auditCount += Add-TempAuditFromSelect `
                -Connection $Connection `
                -TargetTableName $tableName `
                -TargetIdFieldName ([string]$source.TargetIdField) `
                -TempTableName $tempName `
                -SelectIntoSql $selectSql `
                -RuleName "Regen data on not-measured plot" `
                -FieldName $fieldName `
                -Message (Get-NoRegenDataOnInactivePlotStatusMessage)
        }
    }

    return $auditCount
}

function Get-NoRegenDataOnInactivePlotStatusVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName
    )

    $source = Get-NoRegenDataOnInactivePlotStatusSourceDefinition -Connection $Connection -TableName $TableName
    if ($null -eq $source) { return "" }
    if (-not (@($source.Fields) -contains $FieldName)) { return "" }

    $fieldExpression = "$($source.SourceAlias).$(Quote-Name $FieldName)"
    $recordedCondition = Get-NoRegenDataRecordedCondition -FieldExpression $fieldExpression
    $whereParts = New-Object System.Collections.Generic.List[string]
    [void]$whereParts.Add([string]$source.InactivePlotStatusCondition)
    [void]$whereParts.Add($recordedCondition)
    if (-not [string]::IsNullOrWhiteSpace([string]$source.PeriodScopeCondition)) {
        [void]$whereParts.Add([string]$source.PeriodScopeCondition)
    }
    $whereCondition = [string]::Join(" AND ", [string[]]$whereParts.ToArray())

    $selectList = Get-AliasSelectList -Connection $Connection -TableName $TableName -Alias ([string]$source.SourceAlias)
    return @"
SELECT $selectList, $($source.PlotStatusExpression) AS [ResolvedPlotStatus], pm.[PeriodNumber] AS [PlotMeasurementPeriod], r.[RegenMeasKey] AS [ResolvedRegenMeasKey]
$($source.FromSql)
WHERE $whereCondition;
"@
}

function Add-TempAuditFromSelect {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TargetTableName,
        [string]$TargetIdFieldName,
        [string]$TempTableName,
        [string]$SelectIntoSql,
        [string]$RuleName,
        [string]$FieldName,
        [string]$Message
    )

    $tempTable = Quote-Name $TempTableName
    try {
        [void](Invoke-NonQuery -Connection $Connection -Sql $SelectIntoSql)
        $rows = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM $tempTable"
        if ($rows -eq 0) { return 0 }

        $auditSql = @"
INSERT INTO [InventoryCleanAudit]
    ([AuditTime], [TableName], [RuleName], [RecordLabel], [FieldName], [SourceRowId], [ObservedValue], [Message])
SELECT
    Now(),
    $(Sql-Text $TargetTableName),
    $(Sql-Text $RuleName),
    Left(([RecordLabel] & ''), 255),
    $(Sql-Text $FieldName),
    Left(([RowID] & ''), 255),
    Left(([ObservedValue] & ''), 255),
    $(Sql-Text $Message)
FROM $tempTable
"@
        [void](Invoke-NonQuery -Connection $Connection -Sql $auditSql)

        $updateSql = "UPDATE $(Quote-Name $TargetTableName) SET [NeedsReview] = True WHERE $(Quote-Name $TargetIdFieldName) In (SELECT [RowID] FROM $tempTable)"
        [void](Invoke-NonQuery -Connection $Connection -Sql $updateSql)

        return $rows
    }
    finally {
        try {
            [void](Invoke-NonQuery -Connection $Connection -Sql "DROP TABLE $tempTable")
        }
        catch {
        }
    }
}

function Add-TempAuditReportFromSelect {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$AuditTableName,
        [string]$TempTableName,
        [string]$SelectIntoSql,
        [string]$RuleName,
        [string]$FieldName,
        [string]$Message
    )

    $tempTable = Quote-Name $TempTableName
    try {
        [void](Invoke-NonQuery -Connection $Connection -Sql $SelectIntoSql)
        $rows = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM $tempTable"
        if ($rows -eq 0) { return 0 }

        $auditSql = @"
INSERT INTO [InventoryCleanAudit]
    ([AuditTime], [TableName], [RuleName], [RecordLabel], [FieldName], [ObservedValue], [Message])
SELECT
    Now(),
    $(Sql-Text $AuditTableName),
    $(Sql-Text $RuleName),
    Left(([RecordLabel] & ''), 255),
    $(Sql-Text $FieldName),
    Left(([ObservedValue] & ''), 255),
    $(Sql-Text $Message)
FROM $tempTable
"@
        [void](Invoke-NonQuery -Connection $Connection -Sql $auditSql)
        return $rows
    }
    finally {
        try {
            [void](Invoke-NonQuery -Connection $Connection -Sql "DROP TABLE $tempTable")
        }
        catch {
        }
    }
}

function Add-TreeRemeasurementChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return 0 }

    Ensure-NeedsReviewColumn -Connection $Connection -TableName "TreeMeasurements"
    $columns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredField in @("MeasurementID", "TreeKey", "PeriodNumber", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $columns -Name $requiredField)) { return 0 }
    }

    $idbhActiveForReview = (Test-ColumnExists -Columns $columns -Name "IDBH") -and (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "IDBH")
    $heightActiveForReview = (Test-ColumnExists -Columns $columns -Name "TotalHeight") -and (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "TotalHeight")
    $treeHistoryActiveForReview = Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "TreeHistory"
    if (-not $idbhActiveForReview -and -not $heightActiveForReview -and -not $treeHistoryActiveForReview) { return 0 }

    $auditCount = 0
    $idbhJump = Get-ControlDecimalValue -Control $dbhGrowthMax -DefaultValue 100
    $heightJump = Get-ControlDecimalValue -Control $heightGrowthMax -DefaultValue 20
    $idbhJumpSql = $idbhJump.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $heightJumpSql = $heightJump.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $metadataAvailability = Get-TreeMetadataAvailability -Connection $Connection -Tables $tables -MeasurementColumns $columns
    $treeColumnsForLiveStatus = @()
    if ($metadataAvailability.IncludeTreeMetadata) {
        try {
            $treeColumnsForLiveStatus = @(Get-TableColumns -Connection $Connection -TableName "Trees")
        }
        catch {
            $treeColumnsForLiveStatus = @()
        }
    }
    $idbhFromSql = Get-TreePreviousMeasurementFromSql `
        -IncludeTreeMetadata:($metadataAvailability.IncludeTreeMetadata) `
        -IncludePlotMetadata:($metadataAvailability.IncludePlotMetadata)
    $treeAlias = if ($metadataAvailability.IncludeTreeMetadata) { "treeMeta" } else { "" }
    $plotAlias = if ($metadataAvailability.IncludePlotMetadata) { "plotMeta" } else { "" }
    $idbhShrinkCondition = Get-LiveTimberDbhShrinkSqlCondition `
        -CurrentAlias "cur" `
        -PreviousAlias "prev" `
        -TreeAlias $treeAlias `
        -PlotAlias $plotAlias
    $currentPeriodScope = Get-SelectedPeriodAliasCondition -Alias "cur"
    $currentPeriodScopeSql = if ([string]::IsNullOrWhiteSpace($currentPeriodScope)) { "" } else { "`n  AND $currentPeriodScope" }
    $currentOnlyPeriodScope = Get-CurrentMeasurementPeriodAliasCondition -Alias "cur"
    $currentOnlyPeriodScopeSql = if ([string]::IsNullOrWhiteSpace($currentOnlyPeriodScope)) { "" } else { "`n  AND $currentOnlyPeriodScope" }
    $selectedDbhRemeasurementPairCondition = Get-SelectedPreviousCurrentPairCondition -EarlierAlias "prev" -LaterAlias "cur"
    $selectedDbhRemeasurementPairScopeSql = if ([string]::IsNullOrWhiteSpace($selectedDbhRemeasurementPairCondition)) { "" } else { "`n  AND $selectedDbhRemeasurementPairCondition" }

    if ($idbhActiveForReview) {
        $idbhMessage = Add-FieldManualTipToMessage `
            -Message "IDBH increased more than the configured period-to-period threshold. Shrinking DBH on live timber trees is reported separately by the IDBH shrinkage check for the selected current/previous period pair." `
            -TableName "TreeMeasurements" `
            -FieldName "IDBH"
        $idbhTemp = "InventoryCleanChange_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
        $idbhSelect = @"
SELECT
    cur.[MeasurementID] AS [RowID],
    (cur.[TreeKey] & ' P' & cur.[PeriodNumber]) AS [RecordLabel],
    ('Previous P' & prev.[PeriodNumber] & ' IDBH=' & (prev.[IDBH] & '') & '; Current P' & cur.[PeriodNumber] & ' IDBH=' & (cur.[IDBH] & '') & '; IDBHJump=' & ((cur.[IDBH] - prev.[IDBH]) & '')) AS [ObservedValue]
INTO $(Quote-Name $idbhTemp)
$idbhFromSql
WHERE cur.[PeriodNumber] Is Not Null
  $currentOnlyPeriodScopeSql
  $selectedDbhRemeasurementPairScopeSql
  AND cur.[IDBH] Is Not Null
  AND prev.[IDBH] Is Not Null
  AND (cur.[IDBH] - prev.[IDBH]) > $idbhJumpSql
"@
        $auditCount += Add-TempAuditFromSelect `
            -Connection $Connection `
            -TargetTableName "TreeMeasurements" `
            -TargetIdFieldName "MeasurementID" `
            -TempTableName $idbhTemp `
            -SelectIntoSql $idbhSelect `
            -RuleName "IDBH jump check" `
            -FieldName "IDBH" `
            -Message $idbhMessage
    }

    if ($heightActiveForReview) {
        $heightFromSql = Get-TreePreviousMeasurementFromSql `
            -IncludeTreeMetadata:($metadataAvailability.IncludeTreeMetadata) `
            -IncludePlotMetadata:($metadataAvailability.IncludePlotMetadata)
        $heightCurrentLiveCondition = Get-TreeHeightStatusSqlCondition `
            -MeasurementColumns $columns `
            -TreeColumns $treeColumnsForLiveStatus `
            -MeasurementAlias "cur" `
            -TreeAlias $treeAlias `
            -TreeHistoryCodes @("0", "5", "10")
        $heightMessage = Add-FieldManualTipToMessage `
            -Message "TotalHeight changed more than the configured period-to-period threshold or decreased on a live tree. Review for transposed numbers, entry errors, broken-top logic, or true biological change." `
            -TableName "TreeMeasurements" `
            -FieldName "TotalHeight"
        $heightTemp = "InventoryCleanChange_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
        $heightSelect = @"
SELECT
    cur.[MeasurementID] AS [RowID],
    (cur.[TreeKey] & ' P' & cur.[PeriodNumber]) AS [RecordLabel],
    ('Previous P' & prev.[PeriodNumber] & ' TotalHeight=' & (prev.[TotalHeight] & '') & '; Current P' & cur.[PeriodNumber] & ' TotalHeight=' & (cur.[TotalHeight] & '') & '; HeightChange=' & ((cur.[TotalHeight] - prev.[TotalHeight]) & '')) AS [ObservedValue]
INTO $(Quote-Name $heightTemp)
$heightFromSql
WHERE cur.[PeriodNumber] Is Not Null
  $currentPeriodScopeSql
  AND cur.[TotalHeight] Is Not Null
  AND prev.[TotalHeight] Is Not Null
  AND (
        (cur.[TotalHeight] - prev.[TotalHeight]) > $heightJumpSql
        OR (($heightCurrentLiveCondition) AND cur.[TotalHeight] < prev.[TotalHeight])
      )
"@
        $auditCount += Add-TempAuditFromSelect `
            -Connection $Connection `
            -TargetTableName "TreeMeasurements" `
            -TargetIdFieldName "MeasurementID" `
            -TempTableName $heightTemp `
            -SelectIntoSql $heightSelect `
            -RuleName "Height remeasurement check" `
            -FieldName "TotalHeight" `
            -Message $heightMessage
    }

    if ($treeHistoryActiveForReview) {
        $historyMessage = Add-FieldManualTipToMessage `
            -Message "TreeHistory changed in a way that conflicts with the field manual transition guidance for live, new mortality/harvest/thin, old mortality/harvest, or missing trees." `
            -TableName "TreeMeasurements" `
            -FieldName "TreeHistory"
        $historyTemp = "InventoryCleanLogic_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
        $historySelect = @"
SELECT
    cur.[MeasurementID] AS [RowID],
    (cur.[TreeKey] & ' P' & cur.[PeriodNumber]) AS [RecordLabel],
    ('Previous P' & prev.[PeriodNumber] & ' TreeHistory=' & (prev.[TreeHistory] & '') & '; Current P' & cur.[PeriodNumber] & ' TreeHistory=' & (cur.[TreeHistory] & '')) AS [ObservedValue]
INTO $(Quote-Name $historyTemp)
FROM
    ([TreeMeasurements] AS cur
    INNER JOIN (
    SELECT c.[TreeKey], c.[PeriodNumber], Max(p.[PeriodNumber]) AS [PrevPeriodNumber]
    FROM [TreeMeasurements] AS c
    INNER JOIN [TreeMeasurements] AS p
        ON c.[TreeKey] = p.[TreeKey]
       AND p.[PeriodNumber] < c.[PeriodNumber]
    GROUP BY c.[TreeKey], c.[PeriodNumber]
) AS link
        ON cur.[TreeKey] = link.[TreeKey]
       AND cur.[PeriodNumber] = link.[PeriodNumber])
    INNER JOIN [TreeMeasurements] AS prev
    ON prev.[TreeKey] = link.[TreeKey]
   AND prev.[PeriodNumber] = link.[PrevPeriodNumber]
WHERE cur.[PeriodNumber] Is Not Null
  AND cur.[TreeHistory] Is Not Null
  AND prev.[TreeHistory] Is Not Null
  AND (
        (prev.[TreeHistory] In (1, 2, 3, 4, 7, 8) AND cur.[TreeHistory] = 0)
        OR (prev.[TreeHistory] In (0, 10) AND cur.[TreeHistory] In (7, 8))
        OR (prev.[TreeHistory] In (1, 2, 3, 4) AND cur.[TreeHistory] In (1, 2, 3, 4))
        OR (prev.[TreeHistory] In (7, 8) AND cur.[TreeHistory] In (1, 2, 3, 4))
      )
"@
        $auditCount += Add-TempAuditFromSelect `
            -Connection $Connection `
            -TargetTableName "TreeMeasurements" `
            -TargetIdFieldName "MeasurementID" `
            -TempTableName $historyTemp `
            -SelectIntoSql $historySelect `
            -RuleName "TreeHistory transition check" `
            -FieldName "TreeHistory" `
            -Message $historyMessage
    }

    return $auditCount
}

function Add-RmdReferenceChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $auditCount = 0
    $projectPeriodScope = Get-SelectedPeriodAliasCondition -Alias "p"
    $projectPeriodWhere = if ([string]::IsNullOrWhiteSpace($projectPeriodScope)) { "1 = 1" } else { $projectPeriodScope }
    $laterPeriodScope = Get-SelectedPeriodAliasCondition -Alias "later"
    $laterPeriodScopeSql = if ([string]::IsNullOrWhiteSpace($laterPeriodScope)) { "" } else { "`n  AND $laterPeriodScope" }

    if ((Test-TableAvailable -Tables $tables -TableName "Trees") -and
        (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements") -and
        (Test-TableAvailable -Tables $tables -TableName "ProjectMeasurementPeriods")) {

        $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
        $treeMeasColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
        if ((Test-ColumnExists -Columns $treeColumns -Name "TreeID") -and
            (Test-ColumnExists -Columns $treeColumns -Name "PlotNumber") -and
            (Test-ColumnExists -Columns $treeColumns -Name "TreeNumber") -and
            (Test-ColumnExists -Columns $treeMeasColumns -Name "TreeID") -and
            (Test-ColumnExists -Columns $treeMeasColumns -Name "PeriodNumber")) {

            $missingTreesTemp = "InventoryCleanRmd_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $missingTreesSelect = @"
SELECT
    (t.[PlotNumber] & '/' & t.[TreeNumber] & ' P' & p.[PeriodNumber]) AS [RecordLabel],
    ('TreeID=' & (t.[TreeID] & '') & '; missing TreeMeasurements row for period ' & (p.[PeriodNumber] & '')) AS [ObservedValue]
INTO $(Quote-Name $missingTreesTemp)
FROM [Trees] AS t, [ProjectMeasurementPeriods] AS p
WHERE $projectPeriodWhere
  AND NOT EXISTS (
        SELECT 1
        FROM [TreeMeasurements] AS tm
        WHERE tm.[TreeID] = t.[TreeID]
          AND tm.[PeriodNumber] = p.[PeriodNumber]
    )
  AND EXISTS (
        SELECT 1
        FROM [TreeMeasurements] AS tm2
        WHERE tm2.[TreeID] = t.[TreeID]
    )
"@
            $auditCount += Add-TempAuditReportFromSelect `
                -Connection $Connection `
                -AuditTableName "TreeMeasurements" `
                -TempTableName $missingTreesTemp `
                -SelectIntoSql $missingTreesSelect `
                -RuleName "Missing tree measurement" `
                -FieldName "Missing row" `
                -Message "Coworker R cleaning script check: a tree has at least one measurement, but is missing a TreeMeasurements row for another project measurement period."
        }
    }

    if ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
        (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") -and
        (Test-TableAvailable -Tables $tables -TableName "ProjectMeasurementPeriods")) {

        $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
        $plotMeasColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
        if ((Test-ColumnExists -Columns $plotColumns -Name "PlotID") -and
            (Test-ColumnExists -Columns $plotColumns -Name "PlotNumber") -and
            (Test-ColumnExists -Columns $plotMeasColumns -Name "PlotID") -and
            (Test-ColumnExists -Columns $plotMeasColumns -Name "PeriodNumber")) {

            $missingPlotsTemp = "InventoryCleanRmd_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $missingPlotsSelect = @"
SELECT
    (pl.[PlotNumber] & ' P' & p.[PeriodNumber]) AS [RecordLabel],
    ('PlotID=' & (pl.[PlotID] & '') & '; missing PlotMeasurements row for period ' & (p.[PeriodNumber] & '')) AS [ObservedValue]
INTO $(Quote-Name $missingPlotsTemp)
FROM [Plots] AS pl, [ProjectMeasurementPeriods] AS p
WHERE $projectPeriodWhere
  AND NOT EXISTS (
        SELECT 1
        FROM [PlotMeasurements] AS meas
        WHERE meas.[PlotID] = pl.[PlotID]
          AND meas.[PeriodNumber] = p.[PeriodNumber]
    )
  AND EXISTS (
        SELECT 1
        FROM [PlotMeasurements] AS meas2
        WHERE meas2.[PlotID] = pl.[PlotID]
    )
"@
            $auditCount += Add-TempAuditReportFromSelect `
                -Connection $Connection `
                -AuditTableName "PlotMeasurements" `
                -TempTableName $missingPlotsTemp `
                -SelectIntoSql $missingPlotsSelect `
                -RuleName "Missing plot measurement" `
                -FieldName "Missing row" `
                -Message "Coworker R cleaning script check: a plot has at least one measurement, but is missing a PlotMeasurements row for another project measurement period."
        }
    }

    if (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements") {
        $columns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
        $metadataAvailability = Get-TreeMetadataAvailability -Connection $Connection -Tables $tables -MeasurementColumns $columns
        $treeColumnsForLiveStatus = @()
        if ($metadataAvailability.IncludeTreeMetadata) {
            try {
                $treeColumnsForLiveStatus = @(Get-TableColumns -Connection $Connection -TableName "Trees")
            }
            catch {
                $treeColumnsForLiveStatus = @()
            }
        }
        Ensure-NeedsReviewColumn -Connection $Connection -TableName "TreeMeasurements"
        $treeHistoryActiveForReview = Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "TreeHistory"
        $idbhActiveForReview = Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "IDBH"
        $heightActiveForReview = Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "TotalHeight"

        if ($treeHistoryActiveForReview -and
            (Test-ColumnExists -Columns $columns -Name "MeasurementID") -and
            (Test-ColumnExists -Columns $columns -Name "TreeID") -and
            (Test-ColumnExists -Columns $columns -Name "TreeKey") -and
            (Test-ColumnExists -Columns $columns -Name "PeriodNumber") -and
            (Test-ColumnExists -Columns $columns -Name "TreeHistory")) {

            $lazarusMessage = Add-FieldManualTipToMessage `
                -Message "Coworker R cleaning script check: a tree previously coded as mortality, harvest, thinned, or old mortality/harvest was later coded live, missed, or new ingrowth. TreeHistory 9 include/non-include corrections are allowed and are not flagged by this check." `
                -TableName "TreeMeasurements" `
                -FieldName "TreeHistory"
            $lazarusTemp = "InventoryCleanRmd_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $lazarusSelect = @"
SELECT DISTINCT
    later.[MeasurementID] AS [RowID],
    (later.[TreeKey] & ' P' & later.[PeriodNumber]) AS [RecordLabel],
    ('Earlier P' & earlier.[PeriodNumber] & ' TreeHistory=' & (earlier.[TreeHistory] & '') & '; Later P' & later.[PeriodNumber] & ' TreeHistory=' & (later.[TreeHistory] & '')) AS [ObservedValue]
INTO $(Quote-Name $lazarusTemp)
FROM [TreeMeasurements] AS earlier
INNER JOIN [TreeMeasurements] AS later
    ON earlier.[TreeID] = later.[TreeID]
WHERE earlier.[TreeHistory] In (1, 2, 3, 4, 7, 8)
  AND later.[PeriodNumber] > earlier.[PeriodNumber]
  AND later.[TreeHistory] In (0, 5, 10)
"@
            $auditCount += Add-TempAuditFromSelect `
                -Connection $Connection `
                -TargetTableName "TreeMeasurements" `
                -TargetIdFieldName "MeasurementID" `
                -TempTableName $lazarusTemp `
                -SelectIntoSql $lazarusSelect `
                -RuleName "Lazarus tree check" `
                -FieldName "TreeHistory" `
                -Message $lazarusMessage
        }

        if ($idbhActiveForReview -and
            (Test-ColumnExists -Columns $columns -Name "MeasurementID") -and
            (Test-ColumnExists -Columns $columns -Name "TreeID") -and
            (Test-ColumnExists -Columns $columns -Name "TreeKey") -and
            (Test-ColumnExists -Columns $columns -Name "PeriodNumber") -and
            (Test-ColumnExists -Columns $columns -Name "TreeHistory") -and
            (Test-ColumnExists -Columns $columns -Name "IDBH")) {

            $shrinkFromSql = Get-TreeMeasurementPairFromSql `
                -IncludeTreeMetadata:($metadataAvailability.IncludeTreeMetadata) `
                -IncludePlotMetadata:($metadataAvailability.IncludePlotMetadata)
            $treeAlias = if ($metadataAvailability.IncludeTreeMetadata) { "treeMeta" } else { "" }
            $plotAlias = if ($metadataAvailability.IncludePlotMetadata) { "plotMeta" } else { "" }
            $shrinkCondition = Get-LiveTimberDbhShrinkSqlCondition `
                -CurrentAlias "later" `
                -PreviousAlias "earlier" `
                -TreeAlias $treeAlias `
                -PlotAlias $plotAlias
            $selectedDbhPairCondition = Get-SelectedPreviousCurrentPairCondition -EarlierAlias "earlier" -LaterAlias "later"
            $selectedDbhPairScopeSql = if ([string]::IsNullOrWhiteSpace($selectedDbhPairCondition)) { "" } else { "`n  AND $selectedDbhPairCondition" }

            $shrinkMessage = Add-FieldManualTipToMessage `
                -Message "IDBH is smaller in the selected current measurement period than in the selected previous measurement period for a live timber tree. Woodland species codes entered in the options are not flagged for shrinking DBH." `
                -TableName "TreeMeasurements" `
                -FieldName "IDBH"
            $shrinkTemp = "InventoryCleanRmd_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $shrinkSelect = @"
SELECT
    later.[MeasurementID] AS [RowID],
    (later.[TreeKey] & ' P' & later.[PeriodNumber]) AS [RecordLabel],
    ('Earlier P' & earlier.[PeriodNumber] & ' IDBH=' & (earlier.[IDBH] & '') & '; Later P' & later.[PeriodNumber] & ' IDBH=' & (later.[IDBH] & '') & '; Shrinkage=' & ((earlier.[IDBH] - later.[IDBH]) & '')) AS [ObservedValue]
INTO $(Quote-Name $shrinkTemp)
$shrinkFromSql
WHERE earlier.[IDBH] Is Not Null
  AND later.[IDBH] Is Not Null
  AND later.[PeriodNumber] > earlier.[PeriodNumber]
  $laterPeriodScopeSql
  $selectedDbhPairScopeSql
  AND $shrinkCondition
"@
            $auditCount += Add-TempAuditFromSelect `
                -Connection $Connection `
                -TargetTableName "TreeMeasurements" `
                -TargetIdFieldName "MeasurementID" `
                -TempTableName $shrinkTemp `
                -SelectIntoSql $shrinkSelect `
                -RuleName "IDBH shrinkage check" `
                -FieldName "IDBH" `
                -Message $shrinkMessage
        }

        if ($heightActiveForReview -and
            (Test-ColumnExists -Columns $columns -Name "MeasurementID") -and
            (Test-ColumnExists -Columns $columns -Name "TreeID") -and
            (Test-ColumnExists -Columns $columns -Name "TreeKey") -and
            (Test-ColumnExists -Columns $columns -Name "PeriodNumber") -and
            (Test-ColumnExists -Columns $columns -Name "TotalHeight")) {

            $deadHeightShrinkAllowedCondition = ""
            if (Test-ColumnExists -Columns $columns -Name "TreeHistory") {
                $deadHeightShrinkAllowedCondition = Get-DeadTreeHeightShrinkAllowedCondition -EarlierAlias "earlier" -LaterAlias "later"
            }
            $deadHeightShrinkAllowedSql = if ([string]::IsNullOrWhiteSpace($deadHeightShrinkAllowedCondition)) { "" } else { "`n  AND NOT ($deadHeightShrinkAllowedCondition)" }
            $duplicateImmediateHeightShrinkSql = ""
            $heightShrinkFromSql = Get-TreeMeasurementPairFromSql `
                -IncludeTreeMetadata:($metadataAvailability.IncludeTreeMetadata) `
                -IncludePlotMetadata:($metadataAvailability.IncludePlotMetadata)
            $heightShrinkTreeAlias = if ($metadataAvailability.IncludeTreeMetadata) { "treeMeta" } else { "" }
            $heightShrinkLaterLiveCondition = Get-TreeHeightStatusSqlCondition `
                -MeasurementColumns $columns `
                -TreeColumns $treeColumnsForLiveStatus `
                -MeasurementAlias "later" `
                -TreeAlias $heightShrinkTreeAlias `
                -TreeHistoryCodes @("0", "5", "10")
            if (Test-ColumnExists -Columns $columns -Name "TreeHistory") {
                $duplicateImmediateHeightShrinkSql = @"
  AND NOT (
        $heightShrinkLaterLiveCondition
        AND earlier.[PeriodNumber] = (
            SELECT Max(priorHeight.[PeriodNumber])
            FROM [TreeMeasurements] AS priorHeight
            WHERE priorHeight.[TreeID] = later.[TreeID]
              AND priorHeight.[PeriodNumber] < later.[PeriodNumber]
        )
      )
"@
            }
            $heightShrinkMessage = Add-FieldManualTipToMessage `
                -Message "Coworker R cleaning script check: TotalHeight is smaller in a later measurement period than in an earlier period. Review for broken top, measurement method, transposed number, wrong tree, or true field change. Live-to-dead height shrinkage is not flagged when the later TreeHistory is a mortality status selected in Run options as requiring height, because dead trees can lose height through breakage." `
                -TableName "TreeMeasurements" `
                -FieldName "TotalHeight"
            $heightShrinkTemp = "InventoryCleanRmd_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
            $heightShrinkSelect = @"
SELECT
    later.[MeasurementID] AS [RowID],
    (later.[TreeKey] & ' P' & later.[PeriodNumber]) AS [RecordLabel],
    ('Earlier P' & earlier.[PeriodNumber] & ' TotalHeight=' & (earlier.[TotalHeight] & '') & '; Later P' & later.[PeriodNumber] & ' TotalHeight=' & (later.[TotalHeight] & '') & '; Shrinkage=' & ((earlier.[TotalHeight] - later.[TotalHeight]) & '')) AS [ObservedValue]
INTO $(Quote-Name $heightShrinkTemp)
$heightShrinkFromSql
WHERE earlier.[TotalHeight] Is Not Null
  AND later.[TotalHeight] Is Not Null
  AND later.[PeriodNumber] > earlier.[PeriodNumber]
  $laterPeriodScopeSql
  AND later.[TotalHeight] < earlier.[TotalHeight]
  $deadHeightShrinkAllowedSql
  $duplicateImmediateHeightShrinkSql
"@
            $auditCount += Add-TempAuditFromSelect `
                -Connection $Connection `
                -TargetTableName "TreeMeasurements" `
                -TargetIdFieldName "MeasurementID" `
                -TempTableName $heightShrinkTemp `
                -SelectIntoSql $heightShrinkSelect `
                -RuleName "Height shrinkage check" `
                -FieldName "TotalHeight" `
                -Message $heightShrinkMessage
        }

        foreach ($pair in @(
            @{ Problem = "Problem1"; Severity = "Severity1" },
            @{ Problem = "Problem2"; Severity = "Severity2" }
        )) {
            $problemField = [string]$pair.Problem
            $severityField = [string]$pair.Severity
            if (-not (Test-ColumnExists -Columns $columns -Name $problemField) -or
                -not (Test-ColumnExists -Columns $columns -Name $severityField)) {
                continue
            }
            if (-not (Test-AppColumnFieldsActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldNames @($problemField, $severityField))) {
                continue
            }

            $labelExpression = "([TreeKey] & ' P' & [PeriodNumber])"
            $problemObservedExpression = Get-ProblemCodeLabelExpression -ValueExpression "[$problemField]"
            $observedExpression = "('$(($problemField))=' & $problemObservedExpression & '; $(($severityField))=' & ([$severityField] & ''))"
            $message = "Coworker R cleaning script check: problem and severity fields should be filled together. Blank, Null, and 0 mean none; values greater than 0 mean entered."
            $auditCount += Add-ConditionAudit `
                -Connection $Connection `
                -Transaction $null `
                -TableName "TreeMeasurements" `
                -RuleName "Problem severity mismatch" `
                -FieldName "$problemField/$severityField" `
                -ObservedExpression $observedExpression `
                -Message $message `
                -RecordLabelExpression $labelExpression `
                -Condition (Get-ProblemSeverityMismatchCondition -ProblemField $problemField -SeverityField $severityField)
        }

        if ((Test-ColumnExists -Columns $columns -Name "TreeHistory") -and
            (Test-ColumnExists -Columns $columns -Name "Problem1") -and
            (Test-ColumnExists -Columns $columns -Name "Severity1") -and
            (Test-AppColumnFieldsActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldNames @("TreeHistory", "Problem1", "Severity1"))) {

            $labelExpression = "([TreeKey] & ' P' & [PeriodNumber])"
            $problemObservedExpression = Get-ProblemCodeLabelExpression -ValueExpression "[Problem1]"
            $observedExpression = "('TreeHistory=' & ([TreeHistory] & '') & '; Problem1=' & $problemObservedExpression & '; Severity1=' & ([Severity1] & ''))"
            $message = Add-FieldManualTipToMessage `
                -Message "TreeMeasurements.Severity1 must be code 3 when TreeHistory is 2 or 3 and Problem1 is recorded. Blank, Null, and 0 mean no Problem1; values greater than 0 mean a Problem1 code was entered." `
                -TableName "TreeMeasurements" `
                -FieldName "Severity1"
            $auditCount += Add-ConditionAudit `
                -Connection $Connection `
                -Transaction $null `
                -TableName "TreeMeasurements" `
                -RuleName "New mortality Severity1 check" `
                -FieldName "Severity1" `
                -ObservedExpression $observedExpression `
                -Message $message `
                -RecordLabelExpression $labelExpression `
                -Condition (Get-NewMortalityProblem1SeverityCondition)
        }
    }

    return $auditCount
}

function Get-TreeStemCountRecordLabelExpression {
    param(
        [object[]]$Columns,
        [string]$Alias = ""
    )

    $prefix = ""
    if (-not [string]::IsNullOrWhiteSpace($Alias)) {
        $prefix = "$Alias."
    }

    if ((Test-ColumnExists -Columns $Columns -Name "TreeKey") -and
        (Test-ColumnExists -Columns $Columns -Name "PeriodNumber")) {
        return "(" + $prefix + "[TreeKey] & ' P' & " + $prefix + "[PeriodNumber])"
    }
    if ((Test-ColumnExists -Columns $Columns -Name "TreeID") -and
        (Test-ColumnExists -Columns $Columns -Name "PeriodNumber")) {
        return "(" + $prefix + "[TreeID] & ' P' & " + $prefix + "[PeriodNumber])"
    }
    if (Test-ColumnExists -Columns $Columns -Name "MeasurementID") {
        return "('MeasurementID=' & (" + $prefix + "[MeasurementID] & ''))"
    }

    return "'TreeMeasurements record'"
}

function Add-TreeStemCountChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    if (-not (Test-ColumnExists -Columns $Columns -Name "StemCount")) { return 0 }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "StemCount")) { return 0 }

    $stemCountMaximum = Get-StemCountMaximumValue
    $stemCountMaxSql = $stemCountMaximum.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    $baseMessage = "Optional tree StemCount check: entered tree StemCount values must be greater than 0 and no more than $stemCountMaxSql. Timber species may have blank StemCount. Woodland/non-timber species entered in the woodland species box require StemCount."
    if ([string]::IsNullOrWhiteSpace($woodlandList)) {
        $baseMessage += " No woodland species codes were entered, so the required StemCount check for woodland/non-timber trees was not applied."
    }
    $message = Add-FieldManualTipToMessage `
        -Message $baseMessage `
        -TableName "TreeMeasurements" `
        -FieldName "StemCount"

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $treeColumns = @()
    $canJoinSpecies =
        (Test-TableAvailable -Tables $tables -TableName "Trees") -and
        (Test-ColumnExists -Columns $Columns -Name "TreeID")
    if ($canJoinSpecies) {
        $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
        $canJoinSpecies =
            (Test-ColumnExists -Columns $treeColumns -Name "TreeID") -and
            (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")
    }

    $presentInvalidCondition = "[StemCount] Is Not Null AND ([StemCount] <= 0 OR [StemCount] > $stemCountMaxSql)"
    if (-not $canJoinSpecies) {
        return Add-ConditionAudit `
            -Connection $Connection `
            -Transaction $null `
            -TableName "TreeMeasurements" `
            -RuleName "Tree stem count check" `
            -FieldName "StemCount" `
            -ObservedExpression "[StemCount]" `
            -Message ($message + " Species-specific blank checks were skipped because Trees.SpeciesCode could not be joined.") `
            -RecordLabelExpression (Get-TreeStemCountRecordLabelExpression -Columns $Columns) `
            -Condition $presentInvalidCondition
    }

    if (-not (Test-ColumnExists -Columns $Columns -Name "MeasurementID")) {
        return Add-ConditionAudit `
            -Connection $Connection `
            -Transaction $null `
            -TableName "TreeMeasurements" `
            -RuleName "Tree stem count check" `
            -FieldName "StemCount" `
            -ObservedExpression "[StemCount]" `
            -Message ($message + " Species-specific blank checks were skipped because TreeMeasurements.MeasurementID was not available for joined review.") `
            -RecordLabelExpression (Get-TreeStemCountRecordLabelExpression -Columns $Columns) `
            -Condition $presentInvalidCondition
    }

    $aliasedPresentInvalidCondition = "tm.[StemCount] Is Not Null AND (tm.[StemCount] <= 0 OR tm.[StemCount] > $stemCountMaxSql)"
    $conditions = New-Object System.Collections.Generic.List[string]
    [void]$conditions.Add("($aliasedPresentInvalidCondition)")
    if (-not [string]::IsNullOrWhiteSpace($woodlandList)) {
        [void]$conditions.Add("(Trim((treeMeta.[SpeciesCode] & '')) In ($woodlandList) AND tm.[StemCount] Is Null)")
    }
    $condition = "(" + ([string]::Join(" OR ", [string[]]$conditions.ToArray())) + ")"
    $periodScopeCondition = Get-SelectedPeriodAliasCondition -Alias "tm"
    if (-not [string]::IsNullOrWhiteSpace($periodScopeCondition)) {
        $condition = "(($condition) AND ($periodScopeCondition))"
    }
    $recordLabelExpression = Get-TreeStemCountRecordLabelExpression -Columns $Columns -Alias "tm"
    $treeStemTemp = "InventoryCleanTreeStem_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $treeStemSelect = @"
SELECT
    tm.[MeasurementID] AS [RowID],
    $recordLabelExpression AS [RecordLabel],
    ('SpeciesCode=' & (treeMeta.[SpeciesCode] & '') & '; StemCount=' & (tm.[StemCount] & '')) AS [ObservedValue]
INTO $(Quote-Name $treeStemTemp)
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $condition
"@

    return Add-TempAuditFromSelect `
        -Connection $Connection `
        -TargetTableName "TreeMeasurements" `
        -TargetIdFieldName "MeasurementID" `
        -TempTableName $treeStemTemp `
        -SelectIntoSql $treeStemSelect `
        -RuleName "Tree stem count check" `
        -FieldName "StemCount" `
        -Message $message
}

function Add-TreeHeightChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    if (-not (Test-ColumnExists -Columns $Columns -Name "TotalHeight")) { return 0 }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "TotalHeight")) { return 0 }

    $heightMaximum = Get-ControlDecimalValue -Control $heightMax -DefaultValue 150
    $heightMaxSql = $heightMaximum.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $treeColumns = @()
    $canJoinTrees = $false
    if ((Test-TableAvailable -Tables $tables -TableName "Trees") -and
        (Test-ColumnExists -Columns $Columns -Name "TreeID")) {
        try {
            $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
            $canJoinTrees = (Test-ColumnExists -Columns $treeColumns -Name "TreeID")
        }
        catch {
            $treeColumns = @()
            $canJoinTrees = $false
        }
    }

    $minorPlotExpression = Get-MinorPlotValueExpression `
        -MeasurementAlias "tm" `
        -MeasurementColumns $Columns `
        -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
        -TreeColumns $treeColumns
    $minorPlotObserved = if ([string]::IsNullOrWhiteSpace($minorPlotExpression)) { "''" } else { $minorPlotExpression }
    $treeHistoryObserved = if (Test-ColumnExists -Columns $Columns -Name "TreeHistory") { "(tm.[TreeHistory] & '')" } else { "''" }
    $problem1Observed = if (Test-ColumnExists -Columns $Columns -Name "Problem1") { Get-ProblemCodeLabelExpression -ValueExpression "tm.[Problem1]" } else { "''" }
    $problem2Observed = if (Test-ColumnExists -Columns $Columns -Name "Problem2") { Get-ProblemCodeLabelExpression -ValueExpression "tm.[Problem2]" } else { "''" }
    $speciesCodeObserved = if ($canJoinTrees -and (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")) { "(treeMeta.[SpeciesCode] & '')" } else { "''" }

    $message = Add-FieldManualTipToMessage `
        -Message ("Entered TotalHeight values must be greater than zero and no more than the configured maximum. Total Height Protocol: $(Get-TotalHeightProtocolDisplayText). Blank TreeMeasurements.TotalHeight is only flagged for the TreeHistory statuses selected in Run options, and rare species entered for 100% height collection are always checked for live TreeHistory 0, 5, and 10. On older not-yet-crosswalked data, TreeClass 1, 2, 3, 4, and 9 and TreeStatus 1, 2, and 3 are also treated as live for TotalHeight checks. Blank/null TreeHistory and TreeHistory 1, 4, 6, 8, and 9 may have blank height. For timber trees, blank height is treated as correct and entered height is flagged when selected Problem1/Problem2 codes indicate height should be blank: $(Get-ProblemHeightNoHeightCodesText), including TreeHistory 10 ingrowth. Woodland species entered in the woodland species box are skipped for this problem-code height rule." + $(if ((Get-HeightRequiredMinorPlotValues).Count -gt 0) { " Missing-height checks are limited to the minor plots entered in Run options, except rare-species 100% height checks." } else { "" })) `
        -TableName "TreeMeasurements" `
        -FieldName "TotalHeight"

    $presentInvalidCondition = "tm.[TotalHeight] Is Not Null AND (tm.[TotalHeight] <= 0 OR tm.[TotalHeight] > $heightMaxSql)"
    $conditions = New-Object System.Collections.Generic.List[string]
    [void]$conditions.Add("($presentInvalidCondition)")

    if (Test-ColumnExists -Columns $Columns -Name "TreeHistory") {
        $missingRequiredCondition = Get-TreeHeightMissingRequiredCondition `
            -TreeHistoryField "tm.[TreeHistory]" `
            -HeightField "tm.[TotalHeight]" `
            -MinorPlotExpression $minorPlotExpression `
            -MeasurementColumns $Columns `
            -TreeColumns $treeColumns `
            -MeasurementAlias "tm" `
            -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" })
        if (-not [string]::IsNullOrWhiteSpace($missingRequiredCondition)) {
            $problemCodeBlankAllowedCondition = Get-ProblemCodeHeightBlankAllowedCondition `
                -MeasurementColumns $Columns `
                -TreeColumns $treeColumns `
                -MeasurementAlias "tm" `
                -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
                -HeightField "tm.[TotalHeight]"
            if (-not [string]::IsNullOrWhiteSpace($problemCodeBlankAllowedCondition)) {
                $missingRequiredCondition = "(($missingRequiredCondition) AND NOT ($problemCodeBlankAllowedCondition))"
            }
            [void]$conditions.Add("($missingRequiredCondition)")
        }
        elseif ((Get-HeightRequiredMinorPlotValues).Count -gt 0) {
            $message += " The entered minor-plot filter could not be applied because DADA could not find a MinorPlot field on TreeMeasurements or joined Trees."
        }

        if (-not (Test-TotalHeightFullLiveProtocolEnabled)) {
            $rareSpeciesMissingCondition = Get-TreeHeightRareSpeciesMissingCondition `
                -MeasurementColumns $Columns `
                -TreeColumns $treeColumns `
                -MeasurementAlias "tm" `
                -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
                -HeightField "tm.[TotalHeight]"
            if (-not [string]::IsNullOrWhiteSpace($rareSpeciesMissingCondition)) {
                $problemCodeBlankAllowedCondition = Get-ProblemCodeHeightBlankAllowedCondition `
                    -MeasurementColumns $Columns `
                    -TreeColumns $treeColumns `
                    -MeasurementAlias "tm" `
                    -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
                    -HeightField "tm.[TotalHeight]"
                if (-not [string]::IsNullOrWhiteSpace($problemCodeBlankAllowedCondition)) {
                    $rareSpeciesMissingCondition = "(($rareSpeciesMissingCondition) AND NOT ($problemCodeBlankAllowedCondition))"
                }
                [void]$conditions.Add("($rareSpeciesMissingCondition)")
            }
        }
    }

    $problemCodeHeightCondition = Get-ProblemCodeHeightShouldBeBlankCondition `
        -MeasurementColumns $Columns `
        -TreeColumns $treeColumns `
        -MeasurementAlias "tm" `
        -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
        -HeightField "tm.[TotalHeight]"
    if (-not [string]::IsNullOrWhiteSpace($problemCodeHeightCondition)) {
        [void]$conditions.Add("($problemCodeHeightCondition)")
    }

    $condition = "(" + ([string]::Join(" OR ", [string[]]$conditions.ToArray())) + ")"
    $periodScopeCondition = Get-SelectedPeriodAliasCondition -Alias "tm"
    if (-not [string]::IsNullOrWhiteSpace($periodScopeCondition)) {
        $condition = "(($condition) AND ($periodScopeCondition))"
    }
    $recordLabelExpression = Get-TreeStemCountRecordLabelExpression -Columns $Columns -Alias "tm"
    $fromSql = if ($canJoinTrees) {
        @"
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
"@
    }
    else {
        "FROM [TreeMeasurements] AS tm"
    }

    $treeHeightTemp = "InventoryCleanTreeHeight_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $treeHeightSelect = @"
SELECT
    tm.[MeasurementID] AS [RowID],
    $recordLabelExpression AS [RecordLabel],
    ('SpeciesCode=' & $speciesCodeObserved & '; TreeHistory=' & $treeHistoryObserved & '; Problem1=' & $problem1Observed & '; Problem2=' & $problem2Observed & '; MinorPlot=' & ($minorPlotObserved & '') & '; TotalHeight=' & (tm.[TotalHeight] & '')) AS [ObservedValue]
INTO $(Quote-Name $treeHeightTemp)
$fromSql
WHERE $condition
"@

    return Add-TempAuditFromSelect `
        -Connection $Connection `
        -TargetTableName "TreeMeasurements" `
        -TargetIdFieldName "MeasurementID" `
        -TempTableName $treeHeightTemp `
        -SelectIntoSql $treeHeightSelect `
        -RuleName "Height range check" `
        -FieldName "TotalHeight" `
        -Message $message
}

function Add-TreeHeightSubsampleProtocolChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    if (-not (Test-TotalHeightSubsampleProtocolEnabled)) { return 0 }
    if (-not (Test-ColumnExists -Columns $Columns -Name "TotalHeight")) { return 0 }
    if (-not (Test-ColumnExists -Columns $Columns -Name "MeasurementID")) { return 0 }
    foreach ($requiredField in @("TreeID", "PeriodNumber", "TreeHistory", "IDBH")) {
        if (-not (Test-ColumnExists -Columns $Columns -Name $requiredField)) { return 0 }
    }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName "TotalHeight")) { return 0 }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "Trees")) { return 0 }
    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "TreeID")) { return 0 }
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")) { return 0 }

    $plotColumn = Find-CandidateColumn -Columns $treeColumns -Patterns @("^plotnumber$", "^plotid$", "^plotkey$")
    $plotExpression = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "tm.[TreeID]" } else { "treeMeta.$(Quote-Name $plotColumn)" }
    $plotLabel = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "TreeID" } else { $plotColumn }
    $speciesExpression = "Trim((treeMeta.[SpeciesCode] & ''))"
    $minimumIdbhSql = (Get-HeightSubsampleMinimumIdbh).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $minimumCountSql = (Get-HeightSubsampleMinimumCount).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $allAtOrAboveSql = (Get-HeightSubsampleAllAtOrAboveIdbh).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $allAtOrAboveEnabled = Test-HeightSubsampleAllAtOrAboveEnabled
    $allAtFlagExpression = if ($allAtOrAboveEnabled) { "IIf(tm.[IDBH] >= $allAtOrAboveSql, 1, 0)" } else { "0" }
    $dbhClassExpression = if ($allAtOrAboveEnabled) {
        "IIf(tm.[IDBH] >= $allAtOrAboveSql, '>= ' & ($allAtOrAboveSql & ''), ((Int((tm.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & ''))"
    }
    else {
        "((Int((tm.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & '')"
    }
    $requiredCountExpression = if ($allAtOrAboveEnabled) {
        "IIf($allAtFlagExpression = 1, Count(*), IIf(Count(*) < $minimumCountSql, Count(*), $minimumCountSql))"
    }
    else {
        "IIf(Count(*) < $minimumCountSql, Count(*), $minimumCountSql)"
    }

    $conditions = New-Object System.Collections.Generic.List[string]
    [void]$conditions.Add("tm.[IDBH] Is Not Null")
    [void]$conditions.Add("tm.[IDBH] >= $minimumIdbhSql")
    [void]$conditions.Add((Get-TreeHeightStatusSqlCondition -MeasurementColumns $Columns -TreeColumns $treeColumns -MeasurementAlias "tm" -TreeAlias "treeMeta" -TreeHistoryCodes @("0", "5", "10")))
    [void]$conditions.Add("$speciesExpression <> ''")

    $rareSpeciesList = Get-HeightRareSpeciesCodesSqlList
    if (-not [string]::IsNullOrWhiteSpace($rareSpeciesList)) {
        [void]$conditions.Add("$speciesExpression Not In ($rareSpeciesList)")
    }

    $problemCodeExceptionCondition = Get-ProblemCodeHeightExceptionCondition `
        -MeasurementColumns $Columns `
        -TreeColumns $treeColumns `
        -MeasurementAlias "tm" `
        -TreeAlias "treeMeta"
    if (-not [string]::IsNullOrWhiteSpace($problemCodeExceptionCondition)) {
        [void]$conditions.Add("NOT ($problemCodeExceptionCondition)")
    }

    $periodScopeCondition = Get-SelectedPeriodAliasCondition -Alias "tm"
    if (-not [string]::IsNullOrWhiteSpace($periodScopeCondition)) {
        [void]$conditions.Add("($periodScopeCondition)")
    }

    $whereClause = [string]::Join(" AND ", [string[]]$conditions.ToArray())
    $message = Add-FieldManualTipToMessage `
        -Message "TotalHeight subsample protocol review. For each plot, species, and 2-inch IDBH class, DADA checks whether the minimum number of eligible live-tree heights were recorded. TreeClass 1, 2, 3, 4, and 9 and TreeStatus 1, 2, and 3 are treated as live status when old data has not been crosswalked. Rare species entered in Run options are excluded from this grouped subsample check because they are checked as 100% live-tree height requirements. Problem codes selected as no-height exceptions, including legacy 72/74/75 aliases, are excluded from the subsample count." `
        -TableName "TreeMeasurements" `
        -FieldName "TotalHeight"

    $plotLabelPrefix = Sql-Text ($plotLabel + "=")
    $recordLabelExpression = "($plotLabelPrefix & ($plotExpression & '') & '; SpeciesCode=' & ($speciesExpression & '') & '; IDBHClass=' & ($dbhClassExpression & ''))"
    $observedExpression = "('Period=' & (tm.[PeriodNumber] & '') & '; ' & $plotLabelPrefix & ($plotExpression & '') & '; SpeciesCode=' & ($speciesExpression & '') & '; IDBHClass=' & ($dbhClassExpression & '') & '; EligibleTrees=' & (Count(*) & '') & '; HeightsRecorded=' & (Sum(IIf(tm.[TotalHeight] Is Not Null, 1, 0)) & '') & '; RequiredHeights=' & (($requiredCountExpression) & ''))"
    $subsampleTemp = "InventoryCleanHeightSubsample_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $selectSql = @"
SELECT
    Min(tm.[MeasurementID]) AS [RowID],
    $recordLabelExpression AS [RecordLabel],
    $observedExpression AS [ObservedValue]
INTO $(Quote-Name $subsampleTemp)
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $whereClause
GROUP BY tm.[PeriodNumber], $plotExpression, $speciesExpression, $dbhClassExpression, $allAtFlagExpression
HAVING Sum(IIf(tm.[TotalHeight] Is Not Null, 1, 0)) < $requiredCountExpression
"@

    $auditCount = Add-TempAuditFromSelect `
        -Connection $Connection `
        -TargetTableName "TreeMeasurements" `
        -TargetIdFieldName "MeasurementID" `
        -TempTableName $subsampleTemp `
        -SelectIntoSql $selectSql `
        -RuleName "TotalHeight subsample review" `
        -FieldName "TotalHeightSubsample" `
        -Message $message

    $treeNumberColumn = Find-CandidateColumn -Columns $treeColumns -Patterns @("^treenumber$", "^tree_number$", "^tree$")
    if (-not [string]::IsNullOrWhiteSpace($treeNumberColumn)) {
        $treeNumberField = Quote-Name $treeNumberColumn
        $gtPlotExpression = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "gt.[TreeID]" } else { "gt.$(Quote-Name $plotColumn)" }
        $otPlotExpression = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "ot.[TreeID]" } else { "ot.$(Quote-Name $plotColumn)" }
        $plotCorrelationForGroup = "Trim(($gtPlotExpression & '')) = Trim(($plotExpression & ''))"
        $plotCorrelationForOrder = "Trim(($otPlotExpression & '')) = Trim(($plotExpression & ''))"

        $groupAllAtFlagExpression = if ($allAtOrAboveEnabled) { "IIf(g.[IDBH] >= $allAtOrAboveSql, 1, 0)" } else { "0" }
        $orderAllAtFlagExpression = if ($allAtOrAboveEnabled) { "IIf(o.[IDBH] >= $allAtOrAboveSql, 1, 0)" } else { "0" }
        $groupDbhClassExpression = if ($allAtOrAboveEnabled) {
            "IIf(g.[IDBH] >= $allAtOrAboveSql, '>= ' & ($allAtOrAboveSql & ''), ((Int((g.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & ''))"
        }
        else {
            "((Int((g.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & '')"
        }
        $orderDbhClassExpression = if ($allAtOrAboveEnabled) {
            "IIf(o.[IDBH] >= $allAtOrAboveSql, '>= ' & ($allAtOrAboveSql & ''), ((Int((o.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & ''))"
        }
        else {
            "((Int((o.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & '')"
        }

        $groupConditions = New-Object System.Collections.Generic.List[string]
        [void]$groupConditions.Add("g.[IDBH] Is Not Null")
        [void]$groupConditions.Add("g.[IDBH] >= $minimumIdbhSql")
        [void]$groupConditions.Add((Get-TreeHeightStatusSqlCondition -MeasurementColumns $Columns -TreeColumns $treeColumns -MeasurementAlias "g" -TreeAlias "gt" -TreeHistoryCodes @("0", "5", "10")))
        [void]$groupConditions.Add("Trim((gt.[SpeciesCode] & '')) <> ''")
        [void]$groupConditions.Add("g.[PeriodNumber] = tm.[PeriodNumber]")
        [void]$groupConditions.Add($plotCorrelationForGroup)
        [void]$groupConditions.Add("Trim((gt.[SpeciesCode] & '')) = $speciesExpression")
        [void]$groupConditions.Add("$groupDbhClassExpression = $dbhClassExpression")
        [void]$groupConditions.Add("$groupAllAtFlagExpression = $allAtFlagExpression")
        if (-not [string]::IsNullOrWhiteSpace($rareSpeciesList)) {
            [void]$groupConditions.Add("Trim((gt.[SpeciesCode] & '')) Not In ($rareSpeciesList)")
        }
        $groupProblemExceptionCondition = Get-ProblemCodeHeightExceptionCondition `
            -MeasurementColumns $Columns `
            -TreeColumns $treeColumns `
            -MeasurementAlias "g" `
            -TreeAlias "gt"
        if (-not [string]::IsNullOrWhiteSpace($groupProblemExceptionCondition)) {
            [void]$groupConditions.Add("NOT ($groupProblemExceptionCondition)")
        }
        $groupWhereClause = [string]::Join(" AND ", [string[]]$groupConditions.ToArray())
        $groupCountSubquery = "(SELECT Count(*) FROM [TreeMeasurements] AS g INNER JOIN [Trees] AS gt ON g.[TreeID] = gt.[TreeID] WHERE $groupWhereClause)"
        $groupHeightCountSubquery = "(SELECT Sum(IIf(g.[TotalHeight] Is Not Null, 1, 0)) FROM [TreeMeasurements] AS g INNER JOIN [Trees] AS gt ON g.[TreeID] = gt.[TreeID] WHERE $groupWhereClause)"
        $requiredForCurrentExpression = if ($allAtOrAboveEnabled) {
            "IIf($allAtFlagExpression = 1, $groupCountSubquery, IIf($groupCountSubquery < $minimumCountSql, $groupCountSubquery, $minimumCountSql))"
        }
        else {
            "IIf($groupCountSubquery < $minimumCountSql, $groupCountSubquery, $minimumCountSql)"
        }

        $orderConditions = New-Object System.Collections.Generic.List[string]
        [void]$orderConditions.Add("o.[IDBH] Is Not Null")
        [void]$orderConditions.Add("o.[IDBH] >= $minimumIdbhSql")
        [void]$orderConditions.Add((Get-TreeHeightStatusSqlCondition -MeasurementColumns $Columns -TreeColumns $treeColumns -MeasurementAlias "o" -TreeAlias "ot" -TreeHistoryCodes @("0", "5", "10")))
        [void]$orderConditions.Add("Trim((ot.[SpeciesCode] & '')) <> ''")
        [void]$orderConditions.Add("o.[PeriodNumber] = tm.[PeriodNumber]")
        [void]$orderConditions.Add($plotCorrelationForOrder)
        [void]$orderConditions.Add("Trim((ot.[SpeciesCode] & '')) = $speciesExpression")
        [void]$orderConditions.Add("$orderDbhClassExpression = $dbhClassExpression")
        [void]$orderConditions.Add("$orderAllAtFlagExpression = $allAtFlagExpression")
        [void]$orderConditions.Add("ot.$treeNumberField Is Not Null")
        [void]$orderConditions.Add("Len(Trim((ot.$treeNumberField & ''))) > 0")
        [void]$orderConditions.Add("(Val(ot.$treeNumberField & '') < Val(treeMeta.$treeNumberField & '') OR (Val(ot.$treeNumberField & '') = Val(treeMeta.$treeNumberField & '') AND o.[MeasurementID] <= tm.[MeasurementID]))")
        if (-not [string]::IsNullOrWhiteSpace($rareSpeciesList)) {
            [void]$orderConditions.Add("Trim((ot.[SpeciesCode] & '')) Not In ($rareSpeciesList)")
        }
        $orderProblemExceptionCondition = Get-ProblemCodeHeightExceptionCondition `
            -MeasurementColumns $Columns `
            -TreeColumns $treeColumns `
            -MeasurementAlias "o" `
            -TreeAlias "ot"
        if (-not [string]::IsNullOrWhiteSpace($orderProblemExceptionCondition)) {
            [void]$orderConditions.Add("NOT ($orderProblemExceptionCondition)")
        }
        $orderWhereClause = [string]::Join(" AND ", [string[]]$orderConditions.ToArray())
        $clockwiseRankSubquery = "(SELECT Count(*) FROM [TreeMeasurements] AS o INNER JOIN [Trees] AS ot ON o.[TreeID] = ot.[TreeID] WHERE $orderWhereClause)"

        $orderCondition = "(($whereClause) AND tm.[TotalHeight] Is Null AND treeMeta.$treeNumberField Is Not Null AND Len(Trim((treeMeta.$treeNumberField & ''))) > 0 AND $groupHeightCountSubquery >= $requiredForCurrentExpression AND $clockwiseRankSubquery <= $requiredForCurrentExpression)"
        $orderMessage = Add-FieldManualTipToMessage `
            -Message "TotalHeight subsample clockwise-order review. DADA assumes TreeNumber starts at north and increases clockwise. Within each plot, species, and 2-inch IDBH class, the first required eligible trees by TreeNumber should have TotalHeight before later TreeNumbers are used to satisfy the subsample." `
            -TableName "TreeMeasurements" `
            -FieldName "TotalHeight"
        $orderObservedExpression = "('Period=' & (tm.[PeriodNumber] & '') & '; ' & $plotLabelPrefix & ($plotExpression & '') & '; SpeciesCode=' & ($speciesExpression & '') & '; IDBHClass=' & ($dbhClassExpression & '') & '; TreeNumber=' & (treeMeta.$treeNumberField & '') & '; ClockwiseRank=' & (($clockwiseRankSubquery) & '') & '; RequiredHeights=' & (($requiredForCurrentExpression) & '') & '; HeightsRecorded=' & (($groupHeightCountSubquery) & ''))"
        $orderRecordLabelExpression = "($plotLabelPrefix & ($plotExpression & '') & '; SpeciesCode=' & ($speciesExpression & '') & '; IDBHClass=' & ($dbhClassExpression & '') & '; TreeNumber=' & (treeMeta.$treeNumberField & ''))"
        $orderTemp = "InventoryCleanHeightOrder_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
        $orderSelect = @"
SELECT
    tm.[MeasurementID] AS [RowID],
    $orderRecordLabelExpression AS [RecordLabel],
    $orderObservedExpression AS [ObservedValue]
INTO $(Quote-Name $orderTemp)
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $orderCondition
"@

        $auditCount += Add-TempAuditFromSelect `
            -Connection $Connection `
            -TargetTableName "TreeMeasurements" `
            -TargetIdFieldName "MeasurementID" `
            -TempTableName $orderTemp `
            -SelectIntoSql $orderSelect `
            -RuleName "TotalHeight subsample order review" `
            -FieldName "TotalHeightSubsample" `
            -Message $orderMessage
    }

    return $auditCount
}

function Get-TreeHeightSubsampleVerificationSql {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredField in @("MeasurementID", "TreeID", "PeriodNumber", "TreeHistory", "IDBH", "TotalHeight")) {
        if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name $requiredField)) { return "" }
    }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "Trees")) { return "" }
    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "TreeID")) { return "" }
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")) { return "" }

    $plotColumn = Find-CandidateColumn -Columns $treeColumns -Patterns @("^plotnumber$", "^plotid$", "^plotkey$")
    $plotExpression = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "tm.[TreeID]" } else { "treeMeta.$(Quote-Name $plotColumn)" }
    $plotAlias = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "TreeID" } else { $plotColumn }
    $speciesExpression = "Trim((treeMeta.[SpeciesCode] & ''))"
    $minimumIdbhSql = (Get-HeightSubsampleMinimumIdbh).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $minimumCountSql = (Get-HeightSubsampleMinimumCount).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $allAtOrAboveSql = (Get-HeightSubsampleAllAtOrAboveIdbh).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $allAtFlagExpression = if (Test-HeightSubsampleAllAtOrAboveEnabled) { "IIf(tm.[IDBH] >= $allAtOrAboveSql, 1, 0)" } else { "0" }
    $dbhClassExpression = if (Test-HeightSubsampleAllAtOrAboveEnabled) {
        "IIf(tm.[IDBH] >= $allAtOrAboveSql, '>= ' & ($allAtOrAboveSql & ''), ((Int((tm.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & ''))"
    }
    else {
        "((Int((tm.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & '')"
    }
    $requiredCountExpression = if (Test-HeightSubsampleAllAtOrAboveEnabled) {
        "IIf($allAtFlagExpression = 1, Count(*), IIf(Count(*) < $minimumCountSql, Count(*), $minimumCountSql))"
    }
    else {
        "IIf(Count(*) < $minimumCountSql, Count(*), $minimumCountSql)"
    }

    $conditions = New-Object System.Collections.Generic.List[string]
    [void]$conditions.Add("tm.[IDBH] Is Not Null")
    [void]$conditions.Add("tm.[IDBH] >= $minimumIdbhSql")
    [void]$conditions.Add((Get-TreeHeightStatusSqlCondition -MeasurementColumns $treeMeasurementColumns -TreeColumns $treeColumns -MeasurementAlias "tm" -TreeAlias "treeMeta" -TreeHistoryCodes @("0", "5", "10")))
    [void]$conditions.Add("$speciesExpression <> ''")
    $rareSpeciesList = Get-HeightRareSpeciesCodesSqlList
    if (-not [string]::IsNullOrWhiteSpace($rareSpeciesList)) {
        [void]$conditions.Add("$speciesExpression Not In ($rareSpeciesList)")
    }
    $problemCodeExceptionCondition = Get-ProblemCodeHeightExceptionCondition `
        -MeasurementColumns $treeMeasurementColumns `
        -TreeColumns $treeColumns `
        -MeasurementAlias "tm" `
        -TreeAlias "treeMeta"
    if (-not [string]::IsNullOrWhiteSpace($problemCodeExceptionCondition)) {
        [void]$conditions.Add("NOT ($problemCodeExceptionCondition)")
    }
    $periodScopeCondition = Get-SelectedPeriodAliasCondition -Alias "tm"
    if (-not [string]::IsNullOrWhiteSpace($periodScopeCondition)) {
        [void]$conditions.Add("($periodScopeCondition)")
    }
    $whereClause = [string]::Join(" AND ", [string[]]$conditions.ToArray())

    return @"
SELECT
    tm.[PeriodNumber],
    $plotExpression AS [$plotAlias],
    $speciesExpression AS [SpeciesCode],
    $dbhClassExpression AS [IDBHClass],
    Count(*) AS [EligibleTrees],
    Sum(IIf(tm.[TotalHeight] Is Not Null, 1, 0)) AS [HeightsRecorded],
    $requiredCountExpression AS [RequiredHeights]
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $whereClause
GROUP BY tm.[PeriodNumber], $plotExpression, $speciesExpression, $dbhClassExpression, $allAtFlagExpression
HAVING Sum(IIf(tm.[TotalHeight] Is Not Null, 1, 0)) < $requiredCountExpression;
"@
}

function Get-TreeHeightSubsampleOrderVerificationSql {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredField in @("MeasurementID", "TreeID", "PeriodNumber", "TreeHistory", "IDBH", "TotalHeight")) {
        if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name $requiredField)) { return "" }
    }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "Trees")) { return "" }
    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    foreach ($requiredField in @("TreeID", "SpeciesCode")) {
        if (-not (Test-ColumnExists -Columns $treeColumns -Name $requiredField)) { return "" }
    }

    $treeNumberColumn = Find-CandidateColumn -Columns $treeColumns -Patterns @("^treenumber$", "^tree_number$", "^tree$")
    if ([string]::IsNullOrWhiteSpace($treeNumberColumn)) { return "" }

    $plotColumn = Find-CandidateColumn -Columns $treeColumns -Patterns @("^plotnumber$", "^plotid$", "^plotkey$")
    $plotExpression = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "tm.[TreeID]" } else { "treeMeta.$(Quote-Name $plotColumn)" }
    $gtPlotExpression = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "gt.[TreeID]" } else { "gt.$(Quote-Name $plotColumn)" }
    $otPlotExpression = if ([string]::IsNullOrWhiteSpace($plotColumn)) { "ot.[TreeID]" } else { "ot.$(Quote-Name $plotColumn)" }
    $treeNumberField = Quote-Name $treeNumberColumn
    $speciesExpression = "Trim((treeMeta.[SpeciesCode] & ''))"
    $minimumIdbhSql = (Get-HeightSubsampleMinimumIdbh).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $minimumCountSql = (Get-HeightSubsampleMinimumCount).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $allAtOrAboveSql = (Get-HeightSubsampleAllAtOrAboveIdbh).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $allAtOrAboveEnabled = Test-HeightSubsampleAllAtOrAboveEnabled
    $allAtFlagExpression = if ($allAtOrAboveEnabled) { "IIf(tm.[IDBH] >= $allAtOrAboveSql, 1, 0)" } else { "0" }
    $dbhClassExpression = if ($allAtOrAboveEnabled) {
        "IIf(tm.[IDBH] >= $allAtOrAboveSql, '>= ' & ($allAtOrAboveSql & ''), ((Int((tm.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & ''))"
    }
    else {
        "((Int((tm.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & '')"
    }
    $rareSpeciesList = Get-HeightRareSpeciesCodesSqlList

    $baseConditions = New-Object System.Collections.Generic.List[string]
    [void]$baseConditions.Add("tm.[IDBH] Is Not Null")
    [void]$baseConditions.Add("tm.[IDBH] >= $minimumIdbhSql")
    [void]$baseConditions.Add((Get-TreeHeightStatusSqlCondition -MeasurementColumns $treeMeasurementColumns -TreeColumns $treeColumns -MeasurementAlias "tm" -TreeAlias "treeMeta" -TreeHistoryCodes @("0", "5", "10")))
    [void]$baseConditions.Add("$speciesExpression <> ''")
    if (-not [string]::IsNullOrWhiteSpace($rareSpeciesList)) {
        [void]$baseConditions.Add("$speciesExpression Not In ($rareSpeciesList)")
    }
    $problemCodeExceptionCondition = Get-ProblemCodeHeightExceptionCondition `
        -MeasurementColumns $treeMeasurementColumns `
        -TreeColumns $treeColumns `
        -MeasurementAlias "tm" `
        -TreeAlias "treeMeta"
    if (-not [string]::IsNullOrWhiteSpace($problemCodeExceptionCondition)) {
        [void]$baseConditions.Add("NOT ($problemCodeExceptionCondition)")
    }
    $periodScopeCondition = Get-SelectedPeriodAliasCondition -Alias "tm"
    if (-not [string]::IsNullOrWhiteSpace($periodScopeCondition)) {
        [void]$baseConditions.Add("($periodScopeCondition)")
    }
    $whereClause = [string]::Join(" AND ", [string[]]$baseConditions.ToArray())

    $groupAllAtFlagExpression = if ($allAtOrAboveEnabled) { "IIf(g.[IDBH] >= $allAtOrAboveSql, 1, 0)" } else { "0" }
    $orderAllAtFlagExpression = if ($allAtOrAboveEnabled) { "IIf(o.[IDBH] >= $allAtOrAboveSql, 1, 0)" } else { "0" }
    $groupDbhClassExpression = if ($allAtOrAboveEnabled) {
        "IIf(g.[IDBH] >= $allAtOrAboveSql, '>= ' & ($allAtOrAboveSql & ''), ((Int((g.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & ''))"
    }
    else {
        "((Int((g.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & '')"
    }
    $orderDbhClassExpression = if ($allAtOrAboveEnabled) {
        "IIf(o.[IDBH] >= $allAtOrAboveSql, '>= ' & ($allAtOrAboveSql & ''), ((Int((o.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & ''))"
    }
    else {
        "((Int((o.[IDBH] - $minimumIdbhSql) / 20) * 2 + 6) & '')"
    }

    $groupConditions = New-Object System.Collections.Generic.List[string]
    [void]$groupConditions.Add("g.[IDBH] Is Not Null")
    [void]$groupConditions.Add("g.[IDBH] >= $minimumIdbhSql")
    [void]$groupConditions.Add((Get-TreeHeightStatusSqlCondition -MeasurementColumns $treeMeasurementColumns -TreeColumns $treeColumns -MeasurementAlias "g" -TreeAlias "gt" -TreeHistoryCodes @("0", "5", "10")))
    [void]$groupConditions.Add("Trim((gt.[SpeciesCode] & '')) <> ''")
    [void]$groupConditions.Add("g.[PeriodNumber] = tm.[PeriodNumber]")
    [void]$groupConditions.Add("Trim(($gtPlotExpression & '')) = Trim(($plotExpression & ''))")
    [void]$groupConditions.Add("Trim((gt.[SpeciesCode] & '')) = $speciesExpression")
    [void]$groupConditions.Add("$groupDbhClassExpression = $dbhClassExpression")
    [void]$groupConditions.Add("$groupAllAtFlagExpression = $allAtFlagExpression")
    if (-not [string]::IsNullOrWhiteSpace($rareSpeciesList)) {
        [void]$groupConditions.Add("Trim((gt.[SpeciesCode] & '')) Not In ($rareSpeciesList)")
    }
    $groupProblemExceptionCondition = Get-ProblemCodeHeightExceptionCondition `
        -MeasurementColumns $treeMeasurementColumns `
        -TreeColumns $treeColumns `
        -MeasurementAlias "g" `
        -TreeAlias "gt"
    if (-not [string]::IsNullOrWhiteSpace($groupProblemExceptionCondition)) {
        [void]$groupConditions.Add("NOT ($groupProblemExceptionCondition)")
    }
    $groupWhereClause = [string]::Join(" AND ", [string[]]$groupConditions.ToArray())
    $groupCountSubquery = "(SELECT Count(*) FROM [TreeMeasurements] AS g INNER JOIN [Trees] AS gt ON g.[TreeID] = gt.[TreeID] WHERE $groupWhereClause)"
    $groupHeightCountSubquery = "(SELECT Sum(IIf(g.[TotalHeight] Is Not Null, 1, 0)) FROM [TreeMeasurements] AS g INNER JOIN [Trees] AS gt ON g.[TreeID] = gt.[TreeID] WHERE $groupWhereClause)"
    $requiredForCurrentExpression = if ($allAtOrAboveEnabled) {
        "IIf($allAtFlagExpression = 1, $groupCountSubquery, IIf($groupCountSubquery < $minimumCountSql, $groupCountSubquery, $minimumCountSql))"
    }
    else {
        "IIf($groupCountSubquery < $minimumCountSql, $groupCountSubquery, $minimumCountSql)"
    }

    $orderConditions = New-Object System.Collections.Generic.List[string]
    [void]$orderConditions.Add("o.[IDBH] Is Not Null")
    [void]$orderConditions.Add("o.[IDBH] >= $minimumIdbhSql")
    [void]$orderConditions.Add((Get-TreeHeightStatusSqlCondition -MeasurementColumns $treeMeasurementColumns -TreeColumns $treeColumns -MeasurementAlias "o" -TreeAlias "ot" -TreeHistoryCodes @("0", "5", "10")))
    [void]$orderConditions.Add("Trim((ot.[SpeciesCode] & '')) <> ''")
    [void]$orderConditions.Add("o.[PeriodNumber] = tm.[PeriodNumber]")
    [void]$orderConditions.Add("Trim(($otPlotExpression & '')) = Trim(($plotExpression & ''))")
    [void]$orderConditions.Add("Trim((ot.[SpeciesCode] & '')) = $speciesExpression")
    [void]$orderConditions.Add("$orderDbhClassExpression = $dbhClassExpression")
    [void]$orderConditions.Add("$orderAllAtFlagExpression = $allAtFlagExpression")
    [void]$orderConditions.Add("ot.$treeNumberField Is Not Null")
    [void]$orderConditions.Add("Len(Trim((ot.$treeNumberField & ''))) > 0")
    [void]$orderConditions.Add("(Val(ot.$treeNumberField & '') < Val(treeMeta.$treeNumberField & '') OR (Val(ot.$treeNumberField & '') = Val(treeMeta.$treeNumberField & '') AND o.[MeasurementID] <= tm.[MeasurementID]))")
    if (-not [string]::IsNullOrWhiteSpace($rareSpeciesList)) {
        [void]$orderConditions.Add("Trim((ot.[SpeciesCode] & '')) Not In ($rareSpeciesList)")
    }
    $orderProblemExceptionCondition = Get-ProblemCodeHeightExceptionCondition `
        -MeasurementColumns $treeMeasurementColumns `
        -TreeColumns $treeColumns `
        -MeasurementAlias "o" `
        -TreeAlias "ot"
    if (-not [string]::IsNullOrWhiteSpace($orderProblemExceptionCondition)) {
        [void]$orderConditions.Add("NOT ($orderProblemExceptionCondition)")
    }
    $orderWhereClause = [string]::Join(" AND ", [string[]]$orderConditions.ToArray())
    $clockwiseRankSubquery = "(SELECT Count(*) FROM [TreeMeasurements] AS o INNER JOIN [Trees] AS ot ON o.[TreeID] = ot.[TreeID] WHERE $orderWhereClause)"
    $orderCondition = "(($whereClause) AND tm.[TotalHeight] Is Null AND treeMeta.$treeNumberField Is Not Null AND Len(Trim((treeMeta.$treeNumberField & ''))) > 0 AND $groupHeightCountSubquery >= $requiredForCurrentExpression AND $clockwiseRankSubquery <= $requiredForCurrentExpression)"
    $tmSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "tm"

    return @"
SELECT
    $tmSelect,
    $plotExpression AS [SubsamplePlot],
    treeMeta.[SpeciesCode] AS [TreeSpeciesCode],
    treeMeta.$treeNumberField AS [TreeNumber],
    $dbhClassExpression AS [IDBHClass],
    $clockwiseRankSubquery AS [ClockwiseRank],
    $requiredForCurrentExpression AS [RequiredHeights],
    $groupHeightCountSubquery AS [HeightsRecorded]
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $orderCondition;
"@
}

function Get-TreeHeightVerificationSql {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $heightMaxSql = (Get-ControlDecimalValue -Control $heightMax -DefaultValue 150).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    $treeColumns = @()
    $canJoinTrees = $false
    if ((Test-TableAvailable -Tables $tables -TableName "Trees") -and
        (Test-ColumnExists -Columns $treeMeasurementColumns -Name "TreeID")) {
        try {
            $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
            $canJoinTrees = (Test-ColumnExists -Columns $treeColumns -Name "TreeID")
        }
        catch {
            $treeColumns = @()
            $canJoinTrees = $false
        }
    }

    $minorPlotExpression = Get-MinorPlotValueExpression `
        -MeasurementAlias "tm" `
        -MeasurementColumns $treeMeasurementColumns `
        -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
        -TreeColumns $treeColumns
    $conditions = New-Object System.Collections.Generic.List[string]
    [void]$conditions.Add("(tm.[TotalHeight] Is Not Null AND (tm.[TotalHeight] <= 0 OR tm.[TotalHeight] > $heightMaxSql))")

    if (Test-ColumnExists -Columns $treeMeasurementColumns -Name "TreeHistory") {
        $missingRequiredCondition = Get-TreeHeightMissingRequiredCondition `
            -TreeHistoryField "tm.[TreeHistory]" `
            -HeightField "tm.[TotalHeight]" `
            -MinorPlotExpression $minorPlotExpression `
            -MeasurementColumns $treeMeasurementColumns `
            -TreeColumns $treeColumns `
            -MeasurementAlias "tm" `
            -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" })
        if (-not [string]::IsNullOrWhiteSpace($missingRequiredCondition)) {
            $problemCodeBlankAllowedCondition = Get-ProblemCodeHeightBlankAllowedCondition `
                -MeasurementColumns $treeMeasurementColumns `
                -TreeColumns $treeColumns `
                -MeasurementAlias "tm" `
                -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
                -HeightField "tm.[TotalHeight]"
            if (-not [string]::IsNullOrWhiteSpace($problemCodeBlankAllowedCondition)) {
                $missingRequiredCondition = "(($missingRequiredCondition) AND NOT ($problemCodeBlankAllowedCondition))"
            }
            [void]$conditions.Add("($missingRequiredCondition)")
        }

        if (-not (Test-TotalHeightFullLiveProtocolEnabled)) {
            $rareSpeciesMissingCondition = Get-TreeHeightRareSpeciesMissingCondition `
                -MeasurementColumns $treeMeasurementColumns `
                -TreeColumns $treeColumns `
                -MeasurementAlias "tm" `
                -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
                -HeightField "tm.[TotalHeight]"
            if (-not [string]::IsNullOrWhiteSpace($rareSpeciesMissingCondition)) {
                $problemCodeBlankAllowedCondition = Get-ProblemCodeHeightBlankAllowedCondition `
                    -MeasurementColumns $treeMeasurementColumns `
                    -TreeColumns $treeColumns `
                    -MeasurementAlias "tm" `
                    -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
                    -HeightField "tm.[TotalHeight]"
                if (-not [string]::IsNullOrWhiteSpace($problemCodeBlankAllowedCondition)) {
                    $rareSpeciesMissingCondition = "(($rareSpeciesMissingCondition) AND NOT ($problemCodeBlankAllowedCondition))"
                }
                [void]$conditions.Add("($rareSpeciesMissingCondition)")
            }
        }
    }

    $problemCodeHeightCondition = Get-ProblemCodeHeightShouldBeBlankCondition `
        -MeasurementColumns $treeMeasurementColumns `
        -TreeColumns $treeColumns `
        -MeasurementAlias "tm" `
        -TreeAlias $(if ($canJoinTrees) { "treeMeta" } else { "" }) `
        -HeightField "tm.[TotalHeight]"
    if (-not [string]::IsNullOrWhiteSpace($problemCodeHeightCondition)) {
        [void]$conditions.Add("($problemCodeHeightCondition)")
    }

    $whereClause = [string]::Join(" OR ", [string[]]$conditions.ToArray())
    $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "tm"
    $tmSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "tm"
    if ($canJoinTrees) {
        $treeMinorSelect = if (-not [string]::IsNullOrWhiteSpace((Get-MinorPlotColumnName -Columns $treeColumns))) { ", treeMeta.$(Quote-Name (Get-MinorPlotColumnName -Columns $treeColumns)) AS [TreeMinorPlot]" } else { "" }
        return @"
SELECT $tmSelect$treeMinorSelect
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $whereClause;
"@
    }

    return "SELECT $tmSelect FROM [TreeMeasurements] AS tm WHERE $whereClause;"
}

function Get-RegenMinorPlotRuleDefinitions {
    return @(
        [pscustomobject]@{
            Label = "timber seedlings"
            SpeciesGroup = "Timber"
            IdbhCodes = @("0")
            MinorPlots = @(Get-RegenTimberSeedlingMinorPlotValues)
        },
        [pscustomobject]@{
            Label = "timber saplings IDBH 20"
            SpeciesGroup = "Timber"
            IdbhCodes = @("20")
            MinorPlots = @(Get-RegenTimberSapling20MinorPlotValues)
        },
        [pscustomobject]@{
            Label = "timber saplings IDBH 40"
            SpeciesGroup = "Timber"
            IdbhCodes = @("40")
            MinorPlots = @(Get-RegenTimberSapling40MinorPlotValues)
        },
        [pscustomobject]@{
            Label = "woodland seedlings"
            SpeciesGroup = "Woodland"
            IdbhCodes = @("0")
            MinorPlots = @(Get-RegenWoodlandSeedlingMinorPlotValues)
        },
        [pscustomobject]@{
            Label = "woodland saplings IDBH 20"
            SpeciesGroup = "Woodland"
            IdbhCodes = @("20")
            MinorPlots = @(Get-RegenWoodlandSapling20MinorPlotValues)
        },
        [pscustomobject]@{
            Label = "woodland saplings IDBH 40"
            SpeciesGroup = "Woodland"
            IdbhCodes = @("40")
            MinorPlots = @(Get-RegenWoodlandSapling40MinorPlotValues)
        }
    )
}

function Add-RuleCoverageDiagnostic {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName,
        [string]$ObservedValue,
        [string]$Message
    )

    Add-AuditRow `
        -Connection $Connection `
        -TableName $TableName `
        -RuleName "Rule coverage diagnostic" `
        -RecordLabel "Coverage diagnostic" `
        -FieldName $FieldName `
        -ObservedValue $ObservedValue `
        -Message $Message

    return 1
}

function Get-RegenSpeciesGroupSqlCondition {
    param([string]$SpeciesGroup)

    $speciesValue = "Trim(([SpeciesCode] & ''))"
    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    if ($SpeciesGroup -eq "Woodland") {
        if ([string]::IsNullOrWhiteSpace($woodlandList)) { return "" }
        return "($speciesValue <> '' AND $speciesValue In ($woodlandList))"
    }

    if ([string]::IsNullOrWhiteSpace($woodlandList)) {
        return "($speciesValue <> '')"
    }

    return "($speciesValue <> '' AND $speciesValue Not In ($woodlandList))"
}

function Get-ScopedRegenDiagnosticCount {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$FieldName,
        [string]$Condition
    )

    if ([string]::IsNullOrWhiteSpace($Condition)) { return 0 }
    $where = Add-PeriodScopeToCondition -Connection $Connection -TableName "RegenMeasurements" -RuleName "Rule coverage diagnostic" -FieldName $FieldName -Condition $Condition
    return Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [RegenMeasurements] WHERE $where"
}

function Get-RegenMinorPlotMismatchCondition {
    param([object]$Rule)

    $minorPlotList = Convert-CodeValuesToSqlList -Values ([string[]]$Rule.MinorPlots)
    if ([string]::IsNullOrWhiteSpace($minorPlotList)) { return "" }

    $speciesCondition = Get-RegenSpeciesGroupSqlCondition -SpeciesGroup ([string]$Rule.SpeciesGroup)
    if ([string]::IsNullOrWhiteSpace($speciesCondition)) { return "" }

    $idbhList = Convert-CodeValuesToSqlList -Values ([string[]]$Rule.IdbhCodes)
    $idbhValue = "Trim(([IDBH] & ''))"
    $minorPlotValue = "Trim(([MinorPlot] & ''))"
    return "($speciesCondition AND $idbhValue In ($idbhList) AND $minorPlotValue <> '' AND $minorPlotValue Not In ($minorPlotList))"
}

function Add-RegenRuleCoverageDiagnostics {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    $auditCount = 0
    $totalRows = 0
    try {
        $totalRows = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [RegenMeasurements]"
    }
    catch {
        return 0
    }

    if ($totalRows -eq 0) { return 0 }

    foreach ($fieldName in @("SpeciesCode", "IDBH", "StemCount", "MinorPlot")) {
        if (Test-ColumnExists -Columns $Columns -Name $fieldName) { continue }
        $auditCount += Add-RuleCoverageDiagnostic `
            -Connection $Connection `
            -TableName "RegenMeasurements" `
            -FieldName $fieldName `
            -ObservedValue "Diagnostic=Required field missing; Field=$fieldName; RegenRows=$totalRows" `
            -Message "DADA found regen records, but RegenMeasurements.$fieldName is missing. Related regen cleaning checks may not fully run until the database template includes this field."
    }

    $metadata = Get-AppColumnReviewMetadata -Connection $Connection
    $speciesKey = Normalize-AppColumnReviewKey -TableName "RegenMeasurements" -FieldName "SpeciesCode"
    if ((Test-ColumnExists -Columns $Columns -Name "SpeciesCode") -and
        [bool]$metadata.HasMetadata -and
        $metadata.ActiveMap.ContainsKey($speciesKey) -and
        -not [bool]$metadata.ActiveMap[$speciesKey]) {
        $dataCondition = Get-RegenSpeciesRequiredCondition -Columns $Columns -SpeciesField "[SpeciesCode]"
        if ([string]::IsNullOrWhiteSpace($dataCondition)) {
            $dataCondition = "1 = 1"
        }
        $scopedCount = Get-ScopedRegenDiagnosticCount -Connection $Connection -FieldName "SpeciesCode" -Condition $dataCondition
        if ($scopedCount -gt 0) {
            $auditCount += Add-RuleCoverageDiagnostic `
                -Connection $Connection `
                -TableName "RegenMeasurements" `
                -FieldName "SpeciesCode" `
                -ObservedValue "Diagnostic=SpeciesCode inactive in AppColumns; MatchingRows=$scopedCount" `
                -Message "RegenMeasurements.SpeciesCode is inactive in AppColumns, so missing/invalid regen species checks may be skipped even when regen data is present. Turn SpeciesCode active in AppColumns if species should be reviewed."
        }
    }

    if ((Test-ColumnExists -Columns $Columns -Name "StemCount") -and -not (Test-ColumnExists -Columns $Columns -Name "SpeciesCode")) {
        $stemDataCount = Get-ScopedRegenDiagnosticCount -Connection $Connection -FieldName "StemCount" -Condition (Get-FieldHasRecordedValueCondition -FieldExpression "[StemCount]")
        if ($stemDataCount -gt 0) {
            $auditCount += Add-RuleCoverageDiagnostic `
                -Connection $Connection `
                -TableName "RegenMeasurements" `
                -FieldName "StemCount" `
                -ObservedValue "Diagnostic=StemCount check lacks SpeciesCode; MatchingRows=$stemDataCount" `
                -Message "DADA found entered regen StemCount values, but RegenMeasurements.SpeciesCode is missing. StemCount range and required-stem checks use SpeciesCode context, so this should be reviewed as a coverage gap."
        }
    }

    $hasMinorPlotPrereqs = (Test-ColumnExists -Columns $Columns -Name "SpeciesCode") -and
        (Test-ColumnExists -Columns $Columns -Name "IDBH") -and
        (Test-ColumnExists -Columns $Columns -Name "MinorPlot")
    if ($hasMinorPlotPrereqs) {
        $activeRegenCondition = Get-RegenSpeciesRecordedCondition -SpeciesField "[SpeciesCode]"
        $activeRegenCount = Get-ScopedRegenDiagnosticCount -Connection $Connection -FieldName "MinorPlot" -Condition $activeRegenCondition
        $rules = @(Get-RegenMinorPlotRuleDefinitions)
        $configuredRules = @($rules | Where-Object { @($_.MinorPlots).Count -gt 0 })
        if ($activeRegenCount -gt 0 -and $configuredRules.Count -eq 0) {
            $auditCount += Add-RuleCoverageDiagnostic `
                -Connection $Connection `
                -TableName "RegenMeasurements" `
                -FieldName "MinorPlot" `
                -ObservedValue "Diagnostic=No regen MinorPlot allowed lists entered; MatchingRows=$activeRegenCount" `
                -Message "DADA will flag blank Regen MinorPlot values, but no regen MinorPlot allowed lists were entered. Nonblank fake MinorPlot values cannot be verified against project-specific seedling/sapling rules until the allowed minor-plot boxes are filled in."
        }

        if ($configuredRules.Count -gt 0) {
            foreach ($rule in $rules) {
                if (@($rule.MinorPlots).Count -gt 0) { continue }
                $speciesCondition = Get-RegenSpeciesGroupSqlCondition -SpeciesGroup ([string]$rule.SpeciesGroup)
                if ([string]::IsNullOrWhiteSpace($speciesCondition)) { continue }
                $idbhList = Convert-CodeValuesToSqlList -Values ([string[]]$rule.IdbhCodes)
                if ([string]::IsNullOrWhiteSpace($idbhList)) { continue }
                $condition = "($speciesCondition AND Trim(([IDBH] & '')) In ($idbhList))"
                $matchingRows = Get-ScopedRegenDiagnosticCount -Connection $Connection -FieldName "MinorPlot" -Condition $condition
                if ($matchingRows -le 0) { continue }

                $auditCount += Add-RuleCoverageDiagnostic `
                    -Connection $Connection `
                    -TableName "RegenMeasurements" `
                    -FieldName "MinorPlot" `
                    -ObservedValue "Diagnostic=MinorPlot allowed list blank; Group=$($rule.Label); MatchingRows=$matchingRows" `
                    -Message "DADA found regen rows for $($rule.Label), but the allowed MinorPlot list for that group is blank. Blank MinorPlot is still flagged, but nonblank fake values for this group cannot be verified until the allowed list is entered."
            }
        }

        $woodlandRulesConfigured = @($rules | Where-Object { ([string]$_.SpeciesGroup -eq "Woodland") -and @($_.MinorPlots).Count -gt 0 })
        if ($woodlandRulesConfigured.Count -gt 0 -and [string]::IsNullOrWhiteSpace((Get-WoodlandSpeciesCodesSqlList))) {
            $auditCount += Add-RuleCoverageDiagnostic `
                -Connection $Connection `
                -TableName "RegenMeasurements" `
                -FieldName "MinorPlot" `
                -ObservedValue "Diagnostic=Woodland MinorPlot rules entered without woodland SpeciesCode list" `
                -Message "Woodland regen MinorPlot rules were entered, but the woodland species list is blank. DADA cannot classify regen rows as woodland for those MinorPlot rules until woodland SpeciesCode values are entered."
        }
    }

    return $auditCount
}

function Get-AppColumnCoverageState {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName
    )

    $metadata = Get-AppColumnReviewMetadata -Connection $Connection
    $key = Normalize-AppColumnReviewKey -TableName $TableName -FieldName $FieldName
    if (-not [bool]$metadata.HasMetadata) {
        return [pscustomobject]@{
            HasMetadata = $false
            HasEntry = $false
            IsActive = $true
            State = "No AppColumns metadata"
        }
    }

    if (-not $metadata.ActiveMap.ContainsKey($key)) {
        return [pscustomobject]@{
            HasMetadata = $true
            HasEntry = $false
            IsActive = $false
            State = "Missing AppColumns entry"
        }
    }

    $isActive = [bool]$metadata.ActiveMap[$key]
    return [pscustomobject]@{
        HasMetadata = $true
        HasEntry = $true
        IsActive = $isActive
        State = $(if ($isActive) { "Active" } else { "Inactive in AppColumns" })
    }
}

function Add-AppColumnCoverageDiagnosticIfNeeded {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName,
        [string]$DiagnosticName,
        [int]$MatchingRows,
        [string]$CheckDescription
    )

    if ($MatchingRows -le 0) { return 0 }

    $state = Get-AppColumnCoverageState -Connection $Connection -TableName $TableName -FieldName $FieldName
    if (-not [bool]$state.HasMetadata -or [bool]$state.IsActive) { return 0 }

    return Add-RuleCoverageDiagnostic `
        -Connection $Connection `
        -TableName $TableName `
        -FieldName $FieldName `
        -ObservedValue "Diagnostic=$DiagnosticName; Field=$TableName.$FieldName; AppColumns=$($state.State); MatchingRows=$MatchingRows" `
        -Message "$CheckDescription DADA found matching records, but $TableName.$FieldName is $($state.State). Turn this field active in AppColumns if it should be reviewed."
}

function Get-RuleCoverageTableCount {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName = "",
        [string]$Condition = "",
        [switch]$AllPeriods
    )

    try {
        $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
        if (-not (Test-TableAvailable -Tables $tables -TableName $TableName)) { return 0 }
        $where = $Condition
        if (-not $AllPeriods) {
            $where = Add-PeriodScopeToCondition -Connection $Connection -TableName $TableName -RuleName "Rule coverage diagnostic" -FieldName $FieldName -Condition $where
        }
        $sql = "SELECT Count(*) FROM $(Quote-Name $TableName)"
        if (-not [string]::IsNullOrWhiteSpace($where)) {
            $sql += " WHERE $where"
        }
        return Get-CountValue -Connection $Connection -Sql $sql
    }
    catch {
        return 0
    }
}

function Get-PlotCoverageStatusSource {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements")) { return $null }

    $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
    $plotColumns = @()
    $canJoinPlots = $false
    if ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
        (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotID")) {
        try {
            $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
            $canJoinPlots = (Test-ColumnExists -Columns $plotColumns -Name "PlotID")
        }
        catch {
            $plotColumns = @()
            $canJoinPlots = $false
        }
    }

    $statusCandidates = @("PlotStatus", "PlotStatusCode", "PlotStatusID", "Status", "StatusCode", "StatusID")
    $plotMeasurementStatusColumn = Get-FirstMatchingColumnName -Columns $plotMeasurementColumns -CandidateNames $statusCandidates
    $statusExpression = Get-RequiredPlotStatusExpression `
        -MeasurementAlias "pm" `
        -MeasurementColumns $plotMeasurementColumns `
        -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
        -PlotColumns $plotColumns
    $fromSql = if ($canJoinPlots) {
        "FROM [PlotMeasurements] AS pm LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]"
    }
    else {
        "FROM [PlotMeasurements] AS pm"
    }

    return [pscustomobject]@{
        Tables = $tables
        PlotMeasurementColumns = $plotMeasurementColumns
        PlotColumns = $plotColumns
        CanJoinPlots = $canJoinPlots
        PlotMeasurementStatusColumn = $plotMeasurementStatusColumn
        StatusExpression = $statusExpression
        FromSql = $fromSql
    }
}

function Get-PlotCoverageStatusCount {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object]$Source,
        [string[]]$StatusCodes,
        [switch]$AllPeriods
    )

    if ($null -eq $Source -or [string]::IsNullOrWhiteSpace([string]$Source.StatusExpression)) { return 0 }
    $statusNumbers = New-Object System.Collections.Generic.List[string]
    foreach ($statusCode in @($StatusCodes)) {
        $number = 0
        if ([int]::TryParse(([string]$statusCode), [ref]$number)) {
            [void]$statusNumbers.Add($number.ToString([System.Globalization.CultureInfo]::InvariantCulture))
        }
    }
    if ($statusNumbers.Count -eq 0) { return 0 }
    $statusList = [string]::Join(", ", [string[]]$statusNumbers.ToArray())

    $where = "Val($($Source.StatusExpression)) In ($statusList)"
    if (-not $AllPeriods) {
        $periodScope = Get-SelectedPeriodAliasCondition -Alias "pm"
        if (-not [string]::IsNullOrWhiteSpace($periodScope)) {
            $where = "(($where) AND ($periodScope))"
        }
    }

    try {
        return Get-CountValue -Connection $Connection -Sql "SELECT Count(*) $($Source.FromSql) WHERE $where"
    }
    catch {
        return 0
    }
}

function Get-FirstCoverageFieldMatch {
    param(
        [object[]]$TableColumnSets,
        [object]$Definition
    )

    foreach ($tableInfo in @($TableColumnSets)) {
        $fieldName = Get-FirstRequiredPlotFieldColumnName -Columns ([object[]]$tableInfo.Columns) -Definition $Definition
        if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
        return [pscustomobject]@{
            TableName = [string]$tableInfo.TableName
            FieldName = $fieldName
        }
    }

    return $null
}

function Add-PlotRuleCoverageDiagnostics {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $auditCount = 0
    $source = Get-PlotCoverageStatusSource -Connection $Connection
    if ($null -eq $source) { return 0 }

    $plotMeasurementRows = Get-RuleCoverageTableCount -Connection $Connection -TableName "PlotMeasurements" -FieldName "PlotStatus"
    if ($plotMeasurementRows -le 0) { return 0 }

    if ([string]::IsNullOrWhiteSpace([string]$source.PlotMeasurementStatusColumn)) {
        $auditCount += Add-RuleCoverageDiagnostic `
            -Connection $Connection `
            -TableName "PlotMeasurements" `
            -FieldName "PlotStatus" `
            -ObservedValue "Diagnostic=PlotStatus field missing from PlotMeasurements; MatchingRows=$plotMeasurementRows" `
            -Message "PlotStatus progression and several plot/tree/regen status-dependent checks need PlotMeasurements.PlotStatus. Add or map the PlotStatus field in the template before relying on those checks."
    }
    else {
        $auditCount += Add-AppColumnCoverageDiagnosticIfNeeded `
            -Connection $Connection `
            -TableName "PlotMeasurements" `
            -FieldName ([string]$source.PlotMeasurementStatusColumn) `
            -DiagnosticName "PlotStatus inactive or missing from AppColumns" `
            -MatchingRows $plotMeasurementRows `
            -CheckDescription "PlotStatus progression and status-dependent cleaning rules use PlotMeasurements.$($source.PlotMeasurementStatusColumn)."
    }

    if (-not (Test-ColumnExists -Columns ([object[]]$source.PlotMeasurementColumns) -Name "PlotID")) {
        $auditCount += Add-RuleCoverageDiagnostic `
            -Connection $Connection `
            -TableName "PlotMeasurements" `
            -FieldName "PlotID" `
            -ObservedValue "Diagnostic=PlotID field missing; MatchingRows=$plotMeasurementRows" `
            -Message "PlotID is needed to link plot, tree, and regen records to PlotStatus. Some not-measured/dropped-plot checks may not fully run until PlotMeasurements.PlotID is available."
    }
    if (-not (Test-ColumnExists -Columns ([object[]]$source.PlotMeasurementColumns) -Name "PeriodNumber")) {
        $auditCount += Add-RuleCoverageDiagnostic `
            -Connection $Connection `
            -TableName "PlotMeasurements" `
            -FieldName "PeriodNumber" `
            -ObservedValue "Diagnostic=PeriodNumber field missing; MatchingRows=$plotMeasurementRows" `
            -Message "PeriodNumber is needed to scope plot checks to the selected measurement periods and to match plot status to tree/regen records."
    }

    $measuredPlotRows = Get-PlotCoverageStatusCount -Connection $Connection -Source $source -StatusCodes @("1", "2")
    if ($measuredPlotRows -gt 0) {
        $plotCustomColumns = @()
        if (Test-TableAvailable -Tables ([object[]]$source.Tables) -TableName "PlotCustomMeasurements") {
            try { $plotCustomColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotCustomMeasurements") } catch { $plotCustomColumns = @() }
        }
        $tableColumnSets = @(
            [pscustomobject]@{ TableName = "PlotMeasurements"; Columns = ([object[]]$source.PlotMeasurementColumns) },
            [pscustomobject]@{ TableName = "Plots"; Columns = ([object[]]$source.PlotColumns) },
            [pscustomobject]@{ TableName = "PlotCustomMeasurements"; Columns = ([object[]]$plotCustomColumns) }
        )

        foreach ($definition in @(Get-RequiredPlotStatusFieldDefinitions)) {
            $match = Get-FirstCoverageFieldMatch -TableColumnSets $tableColumnSets -Definition $definition
            $displayName = [string]$definition.DisplayName
            if ($null -eq $match) {
                $auditCount += Add-RuleCoverageDiagnostic `
                    -Connection $Connection `
                    -TableName "PlotMeasurements" `
                    -FieldName $displayName `
                    -ObservedValue "Diagnostic=Required plot field missing from template; Field=$displayName; MatchingRows=$measuredPlotRows" `
                    -Message "$displayName is expected for plots with PlotStatus 1 or 2, but DADA could not find a matching field in Plots, PlotMeasurements, or PlotCustomMeasurements. Add/map the field if this item should be checked."
                continue
            }

            $auditCount += Add-AppColumnCoverageDiagnosticIfNeeded `
                -Connection $Connection `
                -TableName ([string]$match.TableName) `
                -FieldName ([string]$match.FieldName) `
                -DiagnosticName "Required plot field inactive or missing from AppColumns" `
                -MatchingRows $measuredPlotRows `
                -CheckDescription "$displayName is expected for plots with PlotStatus 1 or 2."
        }
    }

    $droppedOtherRows = Get-PlotCoverageStatusCount -Connection $Connection -Source $source -StatusCodes @("7")
    if ($droppedOtherRows -gt 0) {
        $remarksDefinition = [pscustomobject]@{
            DisplayName = "PlotRemarks"
            CandidateNames = @("PlotRemarks", "RemarksPlot", "Remarks", "Remark", "PlotNotes", "PlotNote")
        }
        $remarksMatch = Get-FirstCoverageFieldMatch -TableColumnSets @(
            [pscustomobject]@{ TableName = "PlotMeasurements"; Columns = ([object[]]$source.PlotMeasurementColumns) },
            [pscustomobject]@{ TableName = "Plots"; Columns = ([object[]]$source.PlotColumns) }
        ) -Definition $remarksDefinition
        if ($null -eq $remarksMatch) {
            $auditCount += Add-RuleCoverageDiagnostic `
                -Connection $Connection `
                -TableName "PlotMeasurements" `
                -FieldName "PlotRemarks" `
                -ObservedValue "Diagnostic=PlotRemarks field missing; MatchingRows=$droppedOtherRows" `
                -Message "Plot remarks are required when PlotStatus is 7, but DADA could not find a plot remarks field in PlotMeasurements or Plots."
        }
        else {
            $auditCount += Add-AppColumnCoverageDiagnosticIfNeeded `
                -Connection $Connection `
                -TableName ([string]$remarksMatch.TableName) `
                -FieldName ([string]$remarksMatch.FieldName) `
                -DiagnosticName "PlotRemarks inactive or missing from AppColumns" `
                -MatchingRows $droppedOtherRows `
                -CheckDescription "Plot remarks are required when PlotStatus is 7."
        }
    }

    return $auditCount
}

function Get-TreeCoverageLiveCount {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$TreeMeasurementColumns
    )

    if (-not (Test-ColumnExists -Columns $TreeMeasurementColumns -Name "TreeHistory")) { return 0 }
    $historyList = Convert-CodeValuesToSqlList -Values (Get-RequiredTreeCrownHistoryCodes)
    $where = "Trim((tm.[TreeHistory] & '')) In ($historyList)"
    $where = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $where -Alias "tm"
    try {
        return Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [TreeMeasurements] AS tm WHERE $where"
    }
    catch {
        return 0
    }
}

function Get-TreeCoverageRadialCandidateCount {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$TreeMeasurementColumns,
        [object[]]$TreeColumns
    )

    if (-not (Test-ColumnExists -Columns $TreeMeasurementColumns -Name "TreeHistory")) { return 0 }

    $problemFields = New-Object System.Collections.Generic.List[string]
    if (Test-ColumnExists -Columns $TreeMeasurementColumns -Name "Problem1") { [void]$problemFields.Add("tm.[Problem1]") }
    if (Test-ColumnExists -Columns $TreeMeasurementColumns -Name "Problem2") { [void]$problemFields.Add("tm.[Problem2]") }

    $historyValue = "Trim((tm.[TreeHistory] & ''))"
    $requiredReason = "$historyValue = '5'"
    $problem121Parts = New-Object System.Collections.Generic.List[string]
    foreach ($problemField in @($problemFields.ToArray())) {
        [void]$problem121Parts.Add("Trim(($problemField & '')) = '121'")
    }
    if ($problem121Parts.Count -gt 0) {
        $liveList = Convert-CodeValuesToSqlList -Values @("0", "5", "10")
        $problem121Condition = "(" + ([string]::Join(" OR ", [string[]]$problem121Parts.ToArray())) + ")"
        $requiredReason = "(($requiredReason) OR ($historyValue In ($liveList) AND $problem121Condition))"
    }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $canJoinTrees = (Test-TableAvailable -Tables $tables -TableName "Trees") -and
        (Test-ColumnExists -Columns $TreeMeasurementColumns -Name "TreeID") -and
        (Test-ColumnExists -Columns $TreeColumns -Name "TreeID") -and
        (Test-ColumnExists -Columns $TreeColumns -Name "SpeciesCode")
    $fromSql = "FROM [TreeMeasurements] AS tm"
    $woodlandList = Get-WoodlandSpeciesCodesSqlList
    if ($canJoinTrees) {
        $fromSql = "FROM [TreeMeasurements] AS tm LEFT JOIN [Trees] AS treeMeta ON tm.[TreeID] = treeMeta.[TreeID]"
        if (-not [string]::IsNullOrWhiteSpace($woodlandList)) {
            $requiredReason = "(($requiredReason) AND (Trim((treeMeta.[SpeciesCode] & '')) = '' OR Trim((treeMeta.[SpeciesCode] & '')) Not In ($woodlandList)))"
        }
    }

    $where = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $requiredReason -Alias "tm"
    try {
        return Get-CountValue -Connection $Connection -Sql "SELECT Count(*) $fromSql WHERE $where"
    }
    catch {
        return 0
    }
}

function Add-TreeRuleCoverageDiagnostics {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return 0 }

    $auditCount = 0
    $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    $treeRows = Get-RuleCoverageTableCount -Connection $Connection -TableName "TreeMeasurements" -FieldName "TreeHistory"
    if ($treeRows -le 0) { return 0 }

    foreach ($requiredField in @("TreeID", "PeriodNumber")) {
        if (Test-ColumnExists -Columns $treeMeasurementColumns -Name $requiredField) { continue }
        $auditCount += Add-RuleCoverageDiagnostic `
            -Connection $Connection `
            -TableName "TreeMeasurements" `
            -FieldName $requiredField `
            -ObservedValue "Diagnostic=Tree link field missing; Field=$requiredField; MatchingRows=$treeRows" `
            -Message "$requiredField is needed to link tree measurements to species, plots, or measurement periods. Some tree cleaning checks may not fully run until this field is available."
    }

    if (-not (Test-ColumnExists -Columns $treeMeasurementColumns -Name "TreeHistory")) {
        $auditCount += Add-RuleCoverageDiagnostic `
            -Connection $Connection `
            -TableName "TreeMeasurements" `
            -FieldName "TreeHistory" `
            -ObservedValue "Diagnostic=TreeHistory field missing; MatchingRows=$treeRows" `
            -Message "TreeHistory is needed for DBH, height, crown, radial increment, ingrowth, and transition checks. DADA cannot fully review tree rules until TreeMeasurements.TreeHistory is available."
    }
    else {
        $auditCount += Add-AppColumnCoverageDiagnosticIfNeeded `
            -Connection $Connection `
            -TableName "TreeMeasurements" `
            -FieldName "TreeHistory" `
            -DiagnosticName "TreeHistory inactive or missing from AppColumns" `
            -MatchingRows $treeRows `
            -CheckDescription "TreeHistory drives DBH, height, crown, radial increment, ingrowth, and transition checks."
    }

    $treeColumns = @()
    $canReadTrees = $false
    if (Test-TableAvailable -Tables $tables -TableName "Trees") {
        try {
            $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
            $canReadTrees = $true
        }
        catch {
            $treeColumns = @()
            $canReadTrees = $false
        }
    }

    if (-not $canReadTrees) {
        $auditCount += Add-RuleCoverageDiagnostic `
            -Connection $Connection `
            -TableName "Trees" `
            -FieldName "SpeciesCode" `
            -ObservedValue "Diagnostic=Trees table unavailable; MatchingRows=$treeRows" `
            -Message "DADA could not read the Trees table. Tree species, woodland/timber filtering, tree numbers, and several tree checks may not fully run."
    }
    else {
        foreach ($requiredField in @("TreeID", "SpeciesCode")) {
            if (Test-ColumnExists -Columns $treeColumns -Name $requiredField) { continue }
            $auditCount += Add-RuleCoverageDiagnostic `
                -Connection $Connection `
                -TableName "Trees" `
                -FieldName $requiredField `
                -ObservedValue "Diagnostic=Trees link/species field missing; Field=$requiredField; MatchingRows=$treeRows" `
                -Message "Trees.$requiredField is needed for tree species reporting, woodland/timber filtering, or joining tree measurements back to tree identity."
        }

        if (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode") {
            $auditCount += Add-AppColumnCoverageDiagnosticIfNeeded `
                -Connection $Connection `
                -TableName "Trees" `
                -FieldName "SpeciesCode" `
                -DiagnosticName "Tree species inactive or missing from AppColumns" `
                -MatchingRows $treeRows `
                -CheckDescription "Trees.SpeciesCode is needed for required tree species checks, species export context, and woodland/timber filtering."
        }
    }

    $liveTreeRows = Get-TreeCoverageLiveCount -Connection $Connection -TreeMeasurementColumns $treeMeasurementColumns
    if ($liveTreeRows -gt 0) {
        foreach ($definition in @(Get-RequiredTreeCrownFieldDefinitions)) {
            $fieldName = Get-FirstRequiredPlotFieldColumnName -Columns $treeMeasurementColumns -Definition $definition
            $displayName = [string]$definition.DisplayName
            if ([string]::IsNullOrWhiteSpace($fieldName)) {
                $auditCount += Add-RuleCoverageDiagnostic `
                    -Connection $Connection `
                    -TableName "TreeMeasurements" `
                    -FieldName $displayName `
                    -ObservedValue "Diagnostic=Required crown field missing; Field=$displayName; MatchingRows=$liveTreeRows" `
                    -Message "$displayName is required for live TreeHistory 0, 5, and 10 trees, but DADA could not find a matching TreeMeasurements field."
                continue
            }

            $auditCount += Add-AppColumnCoverageDiagnosticIfNeeded `
                -Connection $Connection `
                -TableName "TreeMeasurements" `
                -FieldName $fieldName `
                -DiagnosticName "Required crown field inactive or missing from AppColumns" `
                -MatchingRows $liveTreeRows `
                -CheckDescription "$displayName is required for live TreeHistory 0, 5, and 10 trees."
        }
    }

    $radialCandidateRows = Get-TreeCoverageRadialCandidateCount -Connection $Connection -TreeMeasurementColumns $treeMeasurementColumns -TreeColumns $treeColumns
    if ($radialCandidateRows -gt 0) {
        $definition = @(Get-RequiredTreeRadialIncrementFieldDefinitions)[0]
        $fieldName = Get-FirstRequiredPlotFieldColumnName -Columns $treeMeasurementColumns -Definition $definition
        if ([string]::IsNullOrWhiteSpace($fieldName)) {
            $auditCount += Add-RuleCoverageDiagnostic `
                -Connection $Connection `
                -TableName "TreeMeasurements" `
                -FieldName "RadialIncrement" `
                -ObservedValue "Diagnostic=Radial increment field missing; MatchingRows=$radialCandidateRows" `
                -Message "Radial increment is required for timber missed trees and timber negative-diameter-growth trees, but DADA could not find a radial increment field in TreeMeasurements."
        }
        else {
            $auditCount += Add-AppColumnCoverageDiagnosticIfNeeded `
                -Connection $Connection `
                -TableName "TreeMeasurements" `
                -FieldName $fieldName `
                -DiagnosticName "Radial increment inactive or missing from AppColumns" `
                -MatchingRows $radialCandidateRows `
                -CheckDescription "Radial increment is required for timber missed trees and timber negative-diameter-growth trees."
        }
    }

    return $auditCount
}

function Add-RegenMinorPlotChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    foreach ($requiredColumn in @("SpeciesCode", "IDBH", "MinorPlot")) {
        if (-not (Test-ColumnExists -Columns $Columns -Name $requiredColumn)) { return 0 }
    }

    $auditCount = 0
    $recordLabel = Get-RecordLabelExpression `
        -PlotField (Find-CandidateColumn -Columns $Columns -Patterns @("^plotnumber$", "^plotkey$", "^plotid$", "^plot$")) `
        -TreeField (Find-CandidateColumn -Columns $Columns -Patterns @("^regenmeaskey$", "^measurementid$", "^minorplot$"))

    $requiredCondition = Get-RegenMinorPlotRequiredCondition -SpeciesField "[SpeciesCode]" -MinorPlotField "[MinorPlot]"
    $requiredMessage = Add-FieldManualTipToMessage `
        -Message "Regen MinorPlot is required when RegenMeasurements.SpeciesCode is greater than 0." `
        -TableName "RegenMeasurements" `
        -FieldName "MinorPlot"
    $requiredObservedExpression = "('SpeciesCode=' & ([SpeciesCode] & '') & '; IDBH=' & ([IDBH] & '') & '; MinorPlot is blank')"

    $auditCount += Add-ConditionAudit `
        -Connection $Connection `
        -Transaction $null `
        -TableName "RegenMeasurements" `
        -RuleName "Required regen minor plot" `
        -FieldName "MinorPlot" `
        -ObservedExpression $requiredObservedExpression `
        -Message $requiredMessage `
        -RecordLabelExpression $recordLabel `
        -Condition $requiredCondition

    foreach ($rule in @(Get-RegenMinorPlotRuleDefinitions)) {
        if (@($rule.MinorPlots).Count -eq 0) { continue }

        $condition = Get-RegenMinorPlotMismatchCondition -Rule $rule
        if ([string]::IsNullOrWhiteSpace($condition)) { continue }

        $allowedText = [string]::Join(", ", [string[]]@($rule.MinorPlots))
        $idbhText = [string]::Join("/", [string[]]@($rule.IdbhCodes))
        $message = Add-FieldManualTipToMessage `
            -Message "Regen MinorPlot does not match the project rule for $($rule.Label) (IDBH $idbhText). Allowed minor plots: $allowedText. Woodland/timber classification uses the woodland species codes entered in Run options." `
            -TableName "RegenMeasurements" `
            -FieldName "MinorPlot"
        $observedExpression = "('SpeciesCode=' & ([SpeciesCode] & '') & '; IDBH=' & ([IDBH] & '') & '; MinorPlot=' & ([MinorPlot] & '') & '; Expected=' & $(Sql-Text $allowedText))"

        $auditCount += Add-ConditionAudit `
            -Connection $Connection `
            -Transaction $null `
            -TableName "RegenMeasurements" `
            -RuleName "Regen minor plot check" `
            -FieldName "MinorPlot" `
            -ObservedExpression $observedExpression `
            -Message $message `
            -RecordLabelExpression $recordLabel `
            -Condition $condition
    }

    return $auditCount
}

function Get-RegenMinorPlotVerificationSql {
    param([System.Data.OleDb.OleDbConnection]$Connection = $null)

    $conditions = New-Object System.Collections.Generic.List[string]
    foreach ($rule in @(Get-RegenMinorPlotRuleDefinitions)) {
        $condition = Get-RegenMinorPlotMismatchCondition -Rule $rule
        if (-not [string]::IsNullOrWhiteSpace($condition)) {
            [void]$conditions.Add($condition)
        }
    }

    if ($conditions.Count -eq 0) {
        return "SELECT * FROM [RegenMeasurements] WHERE 1 = 0;"
    }

    $where = [string]::Join(" OR ", [string[]]$conditions.ToArray())
    if ($null -ne $Connection) {
        $where = Add-PeriodScopeToCondition -Connection $Connection -TableName "RegenMeasurements" -RuleName "Regen minor plot check" -FieldName "MinorPlot" -Condition $where
    }

    return "SELECT * FROM [RegenMeasurements] WHERE $where;"
}

function Get-RequiredRegenMinorPlotVerificationSql {
    param([System.Data.OleDb.OleDbConnection]$Connection = $null)

    $condition = Get-RegenMinorPlotRequiredCondition -SpeciesField "[SpeciesCode]" -MinorPlotField "[MinorPlot]"
    if ($null -ne $Connection) {
        $condition = Add-PeriodScopeToCondition -Connection $Connection -TableName "RegenMeasurements" -RuleName "Required regen minor plot" -FieldName "MinorPlot" -Condition $condition
    }

    return "SELECT * FROM [RegenMeasurements] WHERE $condition;"
}

function Get-FirstRequiredPlotFieldColumnName {
    param(
        [object[]]$Columns,
        [object]$Definition
    )

    foreach ($candidate in @($Definition.CandidateNames)) {
        $candidateNormalized = Normalize-FieldName $candidate
        foreach ($column in $Columns) {
            $name = [string]$column.Name
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            if ((Normalize-FieldName $name) -eq $candidateNormalized) {
                return $name
            }
        }
    }

    return ""
}

function Get-RequiredPlotStatusExpression {
    param(
        [string]$MeasurementAlias = "",
        [object[]]$MeasurementColumns = @(),
        [string]$PlotAlias = "",
        [object[]]$PlotColumns = @()
    )

    $statusCandidates = @("PlotStatus", "PlotStatusCode", "PlotStatusID", "Status", "StatusCode", "StatusID")
    $measurementExpression = ""
    $measurementColumn = Get-FirstMatchingColumnName -Columns $MeasurementColumns -CandidateNames $statusCandidates
    if (-not [string]::IsNullOrWhiteSpace($MeasurementAlias) -and -not [string]::IsNullOrWhiteSpace($measurementColumn)) {
        $measurementExpression = "Trim(($MeasurementAlias.$(Quote-Name $measurementColumn) & ''))"
    }

    $plotExpression = ""
    $plotColumn = Get-FirstMatchingColumnName -Columns $PlotColumns -CandidateNames $statusCandidates
    if (-not [string]::IsNullOrWhiteSpace($PlotAlias) -and -not [string]::IsNullOrWhiteSpace($plotColumn)) {
        $plotExpression = "Trim(($PlotAlias.$(Quote-Name $plotColumn) & ''))"
    }

    if (-not [string]::IsNullOrWhiteSpace($measurementExpression) -and -not [string]::IsNullOrWhiteSpace($plotExpression)) {
        return "IIf(Len($measurementExpression) > 0, $measurementExpression, $plotExpression)"
    }
    if (-not [string]::IsNullOrWhiteSpace($measurementExpression)) { return $measurementExpression }
    if (-not [string]::IsNullOrWhiteSpace($plotExpression)) { return $plotExpression }
    return ""
}

function Get-RequiredPlotRecordLabelExpression {
    param(
        [string]$MeasurementAlias = "",
        [object[]]$MeasurementColumns = @(),
        [string]$PlotAlias = "",
        [object[]]$PlotColumns = @(),
        [string]$FallbackExpression = "'plot record'"
    )

    $plotCandidates = @("PlotNumber", "PlotNo", "Plot", "PlotKey", "PlotID")
    $baseExpression = ""
    $plotColumn = Get-FirstMatchingColumnName -Columns $PlotColumns -CandidateNames $plotCandidates
    if (-not [string]::IsNullOrWhiteSpace($PlotAlias) -and -not [string]::IsNullOrWhiteSpace($plotColumn)) {
        $baseExpression = "($PlotAlias.$(Quote-Name $plotColumn) & '')"
    }

    if ([string]::IsNullOrWhiteSpace($baseExpression)) {
        $measurementPlotColumn = Get-FirstMatchingColumnName -Columns $MeasurementColumns -CandidateNames $plotCandidates
        if (-not [string]::IsNullOrWhiteSpace($MeasurementAlias) -and -not [string]::IsNullOrWhiteSpace($measurementPlotColumn)) {
            $baseExpression = "($MeasurementAlias.$(Quote-Name $measurementPlotColumn) & '')"
        }
    }

    if ([string]::IsNullOrWhiteSpace($baseExpression)) {
        $baseExpression = $FallbackExpression
    }

    $periodColumn = Get-FirstMatchingColumnName -Columns $MeasurementColumns -CandidateNames @("PeriodNumber", "Period")
    if (-not [string]::IsNullOrWhiteSpace($MeasurementAlias) -and -not [string]::IsNullOrWhiteSpace($periodColumn)) {
        return "($baseExpression & ' P' & ($MeasurementAlias.$(Quote-Name $periodColumn) & ''))"
    }

    return $baseExpression
}

function Get-MissingRequiredPlotFieldCondition {
    param(
        [string]$FieldExpression,
        [string]$StatusExpression
    )

    if ([string]::IsNullOrWhiteSpace($FieldExpression) -or [string]::IsNullOrWhiteSpace($StatusExpression)) { return "" }
    return "(Val($StatusExpression) In (1, 2) AND ($FieldExpression Is Null OR Len(Trim(($FieldExpression & ''))) = 0))"
}

function Get-RemarksValueExpression {
    param([object[]]$Sources)

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($source in $Sources) {
        $alias = [string]$source.Alias
        $columns = [object[]]$source.Columns
        foreach ($columnName in @(Get-RemarkColumnNames -Columns $columns)) {
            if ([string]::IsNullOrWhiteSpace($alias)) { continue }
            [void]$parts.Add("($alias.$(Quote-Name $columnName) & '')")
        }
    }

    if ($parts.Count -eq 0) { return "''" }
    return "Trim(" + ([string]::Join(" & ' ' & ", [string[]]$parts.ToArray())) + ")"
}

function Get-MissingPlotStatusRemarkCondition {
    param(
        [string]$StatusExpression,
        [string]$RemarksExpression
    )

    if ([string]::IsNullOrWhiteSpace($StatusExpression) -or [string]::IsNullOrWhiteSpace($RemarksExpression)) { return "" }
    return "(Val($StatusExpression) = 7 AND Len(Trim(($RemarksExpression & ''))) = 0)"
}

function Get-PlotStatusProgressionQueryParts {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [switch]$ForVerification
    )

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements")) { return $null }

    $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements")
    foreach ($requiredField in @("PlotID", "PeriodNumber")) {
        if (-not (Test-ColumnExists -Columns $plotMeasurementColumns -Name $requiredField)) { return $null }
    }

    $targetIdFieldName = Get-FirstMatchingColumnName -Columns $plotMeasurementColumns -CandidateNames @("MeasurementID", "ID")
    if ([string]::IsNullOrWhiteSpace($targetIdFieldName)) { return $null }

    $statusCandidates = @("PlotStatus", "PlotStatusCode", "PlotStatusID", "Status", "StatusCode", "StatusID")
    $statusColumn = Get-FirstMatchingColumnName -Columns $plotMeasurementColumns -CandidateNames $statusCandidates
    if ([string]::IsNullOrWhiteSpace($statusColumn)) { return $null }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "PlotMeasurements" -FieldName $statusColumn)) { return $null }

    $plotColumns = @()
    try {
        if (Test-TableAvailable -Tables $tables -TableName "Plots") {
            $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots")
        }
    }
    catch {
        $plotColumns = @()
    }

    $canJoinPlots = ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
        (Test-ColumnExists -Columns $plotColumns -Name "PlotID"))
    $fromSql = if ($canJoinPlots) {
        @"
FROM ([PlotMeasurements] AS pm
LEFT JOIN [Plots] AS pl
    ON pm.[PlotID] = pl.[PlotID])
LEFT JOIN (
    SELECT firstPm.[PlotID], Min(Val((firstPm.[PeriodNumber] & ''))) AS [FirstInstallPeriod]
    FROM [PlotMeasurements] AS firstPm
    WHERE Trim((firstPm.$(Quote-Name $statusColumn) & '')) <> ''
      AND Val(Trim((firstPm.$(Quote-Name $statusColumn) & ''))) = 2
    GROUP BY firstPm.[PlotID]
) AS firstInstall
    ON pm.[PlotID] = firstInstall.[PlotID]
"@
    }
    else {
        @"
FROM [PlotMeasurements] AS pm
LEFT JOIN (
    SELECT firstPm.[PlotID], Min(Val((firstPm.[PeriodNumber] & ''))) AS [FirstInstallPeriod]
    FROM [PlotMeasurements] AS firstPm
    WHERE Trim((firstPm.$(Quote-Name $statusColumn) & '')) <> ''
      AND Val(Trim((firstPm.$(Quote-Name $statusColumn) & ''))) = 2
    GROUP BY firstPm.[PlotID]
) AS firstInstall
    ON pm.[PlotID] = firstInstall.[PlotID]
"@
    }

    $statusText = "Trim((pm.$(Quote-Name $statusColumn) & ''))"
    $periodValue = "Val((pm.[PeriodNumber] & ''))"
    $firstInstallPeriod = "firstInstall.[FirstInstallPeriod]"
    $hasPeriod = "Trim((pm.[PeriodNumber] & '')) <> ''"
    $beforeInstallCondition = "$statusText <> '' AND ($firstInstallPeriod Is Null OR $periodValue < $firstInstallPeriod OR ($periodValue = $firstInstallPeriod AND Val($statusText) <> 2))"
    $afterInstallCondition = "$firstInstallPeriod Is Not Null AND $periodValue > $firstInstallPeriod AND ($statusText = '' OR Val($statusText) Not In (1, 4, 5, 6, 7))"
    $whereCondition = "$hasPeriod AND (($beforeInstallCondition) OR ($afterInstallCondition))"
    $recordLabel = Get-RequiredPlotRecordLabelExpression `
        -MeasurementAlias "pm" `
        -MeasurementColumns $plotMeasurementColumns `
        -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
        -PlotColumns $plotColumns `
        -FallbackExpression "(pm.$(Quote-Name $targetIdFieldName) & '')"
    $reasonExpression = @"
IIf($firstInstallPeriod Is Null,
    'first nonblank PlotStatus should be 2 before any later status is used',
    IIf($periodValue < $firstInstallPeriod,
        'status was recorded before initial install PlotStatus 2',
        IIf($periodValue = $firstInstallPeriod AND Val($statusText) <> 2,
            'initial install period has a non-2 PlotStatus',
            'after initial install, PlotStatus must be 1, 4, 5, 6, or 7 and cannot be blank')))
"@
    $observedExpression = "('PlotStatus=' & ($statusText & '') & '; FirstInstallPeriod=' & IIf($firstInstallPeriod Is Null, 'none', ($firstInstallPeriod & '')) & '; ' & $reasonExpression)"

    $selectList = if ($ForVerification) {
        $baseSelect = Get-AliasSelectList -Connection $Connection -TableName "PlotMeasurements" -Alias "pm"
        if ($canJoinPlots -and (Test-ColumnExists -Columns $plotColumns -Name "PlotNumber")) {
            "$baseSelect, pl.[PlotNumber] AS [PlotNumber], $firstInstallPeriod AS [FirstInstallPeriod], $observedExpression AS [ProgressionIssue]"
        }
        else {
            "$baseSelect, $firstInstallPeriod AS [FirstInstallPeriod], $observedExpression AS [ProgressionIssue]"
        }
    }
    else {
        @"
pm.$(Quote-Name $targetIdFieldName) AS [RowID],
    $recordLabel AS [RecordLabel],
    $observedExpression AS [ObservedValue]
"@
    }

    return [pscustomobject]@{
        SelectList = $selectList
        FromSql = $fromSql
        WhereCondition = $whereCondition
        TargetIdFieldName = $targetIdFieldName
        StatusColumn = $statusColumn
    }
}

function Get-PlotStatusProgressionMessage {
    return "PlotStatus progression should start with blank/null values until the first measured period. The first nonblank PlotStatus for a plot should be 2 (initial install). After PlotStatus 2, later periods must have PlotStatus 1, 4, 5, 6, or 7 and cannot be blank; PlotStatus 5 should not appear before the plot has ever received PlotStatus 2."
}

function Add-PlotStatusProgressionChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $source = Get-PlotStatusProgressionQueryParts -Connection $Connection
    if ($null -eq $source) { return 0 }

    Ensure-NeedsReviewColumn -Connection $Connection -TableName "PlotMeasurements"
    $tempName = "InventoryCleanPlotStatus_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $selectSql = @"
SELECT
    $($source.SelectList)
INTO $(Quote-Name $tempName)
$($source.FromSql)
WHERE $($source.WhereCondition)
"@

    return Add-TempAuditFromSelect `
        -Connection $Connection `
        -TargetTableName "PlotMeasurements" `
        -TargetIdFieldName ([string]$source.TargetIdFieldName) `
        -TempTableName $tempName `
        -SelectIntoSql $selectSql `
        -RuleName "PlotStatus progression check" `
        -FieldName "PlotStatus" `
        -Message (Get-PlotStatusProgressionMessage)
}

function Get-PlotStatusProgressionVerificationSql {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $source = Get-PlotStatusProgressionQueryParts -Connection $Connection -ForVerification
    if ($null -eq $source) { return "" }

    return @"
SELECT $($source.SelectList)
$($source.FromSql)
WHERE $($source.WhereCondition);
"@
}

function Add-PlotStatusRemarkChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [object[]]$Columns
    )

    if ($TableName -notin @("Plots", "PlotMeasurements")) { return 0 }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $plotColumns = @()
    $plotMeasurementColumns = @()
    try { if (Test-TableAvailable -Tables $tables -TableName "Plots") { $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots") } } catch { $plotColumns = @() }
    try { if (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") { $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements") } } catch { $plotMeasurementColumns = @() }
    $statusCandidates = @("PlotStatus", "PlotStatusCode", "PlotStatusID", "Status", "StatusCode", "StatusID")

    $message = Add-FieldManualTipToMessage `
        -Message "PlotRemarks is required when PlotStatus is code 7." `
        -TableName $TableName `
        -FieldName "PlotRemarks"

    $tempName = "InventoryCleanPlotRemark_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $selectSql = ""
    $targetIdFieldName = ""

    switch ($TableName) {
        "PlotMeasurements" {
            $targetIdFieldName = Get-FirstMatchingColumnName -Columns $Columns -CandidateNames @("MeasurementID", "ID")
            if ([string]::IsNullOrWhiteSpace($targetIdFieldName)) { return 0 }
            if ([string]::IsNullOrWhiteSpace((Get-FirstMatchingColumnName -Columns $Columns -CandidateNames $statusCandidates))) { return 0 }

            $canJoinPlots = ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
                (Test-ColumnExists -Columns $Columns -Name "PlotID") -and
                (Test-ColumnExists -Columns $plotColumns -Name "PlotID"))
            $fromSql = if ($canJoinPlots) {
                "FROM [PlotMeasurements] AS pm LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]"
            }
            else {
                "FROM [PlotMeasurements] AS pm"
            }
            $statusExpression = Get-RequiredPlotStatusExpression `
                -MeasurementAlias "pm" `
                -MeasurementColumns $Columns `
                -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                -PlotColumns $plotColumns
            if ([string]::IsNullOrWhiteSpace($statusExpression)) { return 0 }

            $remarksExpression = Get-RemarksValueExpression -Sources @(
                @{ Alias = "pm"; Columns = $Columns },
                @{ Alias = $(if ($canJoinPlots) { "pl" } else { "" }); Columns = $plotColumns }
            )
            $condition = Get-MissingPlotStatusRemarkCondition -StatusExpression $statusExpression -RemarksExpression $remarksExpression
            $periodScope = Get-SelectedPeriodAliasCondition -Alias "pm"
            if (-not [string]::IsNullOrWhiteSpace($periodScope)) { $condition = "(($condition) AND ($periodScope))" }
            $recordLabel = Get-RequiredPlotRecordLabelExpression `
                -MeasurementAlias "pm" `
                -MeasurementColumns $Columns `
                -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                -PlotColumns $plotColumns `
                -FallbackExpression "(pm.$(Quote-Name $targetIdFieldName) & '')"
            $observedExpression = "('PlotStatus=' & ($statusExpression & '') & '; PlotRemarks is blank')"
            $selectSql = @"
SELECT
    pm.$(Quote-Name $targetIdFieldName) AS [RowID],
    $recordLabel AS [RecordLabel],
    $observedExpression AS [ObservedValue]
INTO $(Quote-Name $tempName)
$fromSql
WHERE $condition
"@
        }
        "Plots" {
            if (-not (Test-ColumnExists -Columns $Columns -Name "PlotID")) { return 0 }
            if ([string]::IsNullOrWhiteSpace((Get-FirstMatchingColumnName -Columns $Columns -CandidateNames $statusCandidates))) { return 0 }
            $targetIdFieldName = "PlotID"

            $canJoinPlotMeasurements = ((Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") -and
                (Test-ColumnExists -Columns $Columns -Name "PlotID") -and
                (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotID"))
            $fromSql = if ($canJoinPlotMeasurements) {
                "FROM [Plots] AS pl INNER JOIN [PlotMeasurements] AS pm ON pl.[PlotID] = pm.[PlotID]"
            }
            else {
                "FROM [Plots] AS pl"
            }
            $statusExpression = Get-RequiredPlotStatusExpression `
                -MeasurementAlias "" `
                -MeasurementColumns @() `
                -PlotAlias "pl" `
                -PlotColumns $Columns
            if ([string]::IsNullOrWhiteSpace($statusExpression)) { return 0 }

            $remarksExpression = Get-RemarksValueExpression -Sources @(
                @{ Alias = "pl"; Columns = $Columns }
            )
            $condition = Get-MissingPlotStatusRemarkCondition -StatusExpression $statusExpression -RemarksExpression $remarksExpression
            if ($canJoinPlotMeasurements) {
                $periodScope = Get-SelectedPeriodAliasCondition -Alias "pm"
                if (-not [string]::IsNullOrWhiteSpace($periodScope)) { $condition = "(($condition) AND ($periodScope))" }
            }
            $recordLabel = Get-RequiredPlotRecordLabelExpression `
                -MeasurementAlias "" `
                -MeasurementColumns @() `
                -PlotAlias "pl" `
                -PlotColumns $Columns `
                -FallbackExpression "(pl.$(Quote-Name $targetIdFieldName) & '')"
            $observedExpression = if ($canJoinPlotMeasurements) {
                "'Selected-period PlotStatus 7; PlotRemarks is blank'"
            }
            else {
                "('PlotStatus=' & ($statusExpression & '') & '; PlotRemarks is blank')"
            }
            $selectSql = @"
SELECT DISTINCT
    pl.$(Quote-Name $targetIdFieldName) AS [RowID],
    $recordLabel AS [RecordLabel],
    $observedExpression AS [ObservedValue]
INTO $(Quote-Name $tempName)
$fromSql
WHERE $condition
"@
        }
    }

    if ([string]::IsNullOrWhiteSpace($selectSql) -or [string]::IsNullOrWhiteSpace($targetIdFieldName)) { return 0 }

    return Add-TempAuditFromSelect `
        -Connection $Connection `
        -TargetTableName $TableName `
        -TargetIdFieldName $targetIdFieldName `
        -TempTableName $tempName `
        -SelectIntoSql $selectSql `
        -RuleName "Required plot remark" `
        -FieldName "PlotRemarks" `
        -Message $message
}

function Add-RequiredPlotStatusFieldChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [object[]]$Columns
    )

    if ($TableName -notin @("Plots", "PlotMeasurements", "PlotCustomMeasurements")) { return 0 }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $plotColumns = @()
    $plotMeasurementColumns = @()
    try { if (Test-TableAvailable -Tables $tables -TableName "Plots") { $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots") } } catch { $plotColumns = @() }
    try { if (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") { $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements") } } catch { $plotMeasurementColumns = @() }

    $auditCount = 0
    $seenFields = @{}
    foreach ($definition in @(Get-RequiredPlotStatusFieldDefinitions)) {
        $fieldName = Get-FirstRequiredPlotFieldColumnName -Columns $Columns -Definition $definition
        if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
        $fieldKey = Normalize-FieldName $fieldName
        if ($seenFields.ContainsKey($fieldKey)) { continue }
        $seenFields[$fieldKey] = $true
        if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName $fieldName)) { continue }

        $displayName = [string]$definition.DisplayName
        $message = Add-FieldManualTipToMessage `
            -Message "$displayName is required for all plots with PlotStatus code 1 or 2." `
            -TableName $TableName `
            -FieldName $fieldName

        $tempName = "InventoryCleanPlotReq_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
        $observedSql = "('PlotStatus=' & (%STATUS% & '') & $(Sql-Text ("; $displayName is blank")))"
        $selectSql = ""
        $targetIdFieldName = ""

        switch ($TableName) {
            "PlotMeasurements" {
                $targetIdFieldName = Get-FirstMatchingColumnName -Columns $Columns -CandidateNames @("MeasurementID", "ID")
                if ([string]::IsNullOrWhiteSpace($targetIdFieldName)) { continue }

                $canJoinPlots = ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
                    (Test-ColumnExists -Columns $Columns -Name "PlotID") -and
                    (Test-ColumnExists -Columns $plotColumns -Name "PlotID"))
                $fromSql = if ($canJoinPlots) {
                    "FROM [PlotMeasurements] AS pm LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]"
                }
                else {
                    "FROM [PlotMeasurements] AS pm"
                }
                $statusExpression = Get-RequiredPlotStatusExpression `
                    -MeasurementAlias "pm" `
                    -MeasurementColumns $Columns `
                    -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                    -PlotColumns $plotColumns
                if ([string]::IsNullOrWhiteSpace($statusExpression)) { continue }

                $condition = Get-MissingRequiredPlotFieldCondition -FieldExpression "pm.$(Quote-Name $fieldName)" -StatusExpression $statusExpression
                $periodScope = Get-SelectedPeriodAliasCondition -Alias "pm"
                if (-not [string]::IsNullOrWhiteSpace($periodScope)) { $condition = "(($condition) AND ($periodScope))" }
                $recordLabel = Get-RequiredPlotRecordLabelExpression `
                    -MeasurementAlias "pm" `
                    -MeasurementColumns $Columns `
                    -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                    -PlotColumns $plotColumns `
                    -FallbackExpression "(pm.$(Quote-Name $targetIdFieldName) & '')"
                $observedExpression = $observedSql.Replace("%STATUS%", $statusExpression)
                $selectSql = @"
SELECT
    pm.$(Quote-Name $targetIdFieldName) AS [RowID],
    $recordLabel AS [RecordLabel],
    $observedExpression AS [ObservedValue]
INTO $(Quote-Name $tempName)
$fromSql
WHERE $condition
"@
            }
            "PlotCustomMeasurements" {
                $targetIdFieldName = Get-FirstMatchingColumnName -Columns $Columns -CandidateNames @("MeasurementID", "ID")
                if ([string]::IsNullOrWhiteSpace($targetIdFieldName)) { continue }
                if (-not ((Test-ColumnExists -Columns $Columns -Name "PlotMeasKey") -and (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotMeasKey"))) { continue }

                $canJoinPlots = ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
                    (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotID") -and
                    (Test-ColumnExists -Columns $plotColumns -Name "PlotID"))
                $fromSql = if ($canJoinPlots) {
                    "FROM ([PlotCustomMeasurements] AS c LEFT JOIN [PlotMeasurements] AS pm ON c.[PlotMeasKey] = pm.[PlotMeasKey]) LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]"
                }
                else {
                    "FROM [PlotCustomMeasurements] AS c LEFT JOIN [PlotMeasurements] AS pm ON c.[PlotMeasKey] = pm.[PlotMeasKey]"
                }
                $statusExpression = Get-RequiredPlotStatusExpression `
                    -MeasurementAlias "pm" `
                    -MeasurementColumns $plotMeasurementColumns `
                    -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                    -PlotColumns $plotColumns
                if ([string]::IsNullOrWhiteSpace($statusExpression)) { continue }

                $condition = Get-MissingRequiredPlotFieldCondition -FieldExpression "c.$(Quote-Name $fieldName)" -StatusExpression $statusExpression
                $periodScope = Get-SelectedPeriodAliasCondition -Alias "pm"
                if (-not [string]::IsNullOrWhiteSpace($periodScope)) { $condition = "(($condition) AND ($periodScope))" }
                $recordLabel = Get-RequiredPlotRecordLabelExpression `
                    -MeasurementAlias "pm" `
                    -MeasurementColumns $plotMeasurementColumns `
                    -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                    -PlotColumns $plotColumns `
                    -FallbackExpression "(c.$(Quote-Name $targetIdFieldName) & '')"
                $observedExpression = $observedSql.Replace("%STATUS%", $statusExpression)
                $selectSql = @"
SELECT
    c.$(Quote-Name $targetIdFieldName) AS [RowID],
    $recordLabel AS [RecordLabel],
    $observedExpression AS [ObservedValue]
INTO $(Quote-Name $tempName)
$fromSql
WHERE $condition
"@
            }
            "Plots" {
                if (-not (Test-ColumnExists -Columns $Columns -Name "PlotID")) { continue }
                $targetIdFieldName = "PlotID"

                $canJoinPlotMeasurements = ((Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") -and
                    (Test-ColumnExists -Columns $Columns -Name "PlotID") -and
                    (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotID"))
                $fromSql = if ($canJoinPlotMeasurements) {
                    "FROM [Plots] AS pl INNER JOIN [PlotMeasurements] AS pm ON pl.[PlotID] = pm.[PlotID]"
                }
                else {
                    "FROM [Plots] AS pl"
                }
                $statusExpression = Get-RequiredPlotStatusExpression `
                    -MeasurementAlias $(if ($canJoinPlotMeasurements) { "pm" } else { "" }) `
                    -MeasurementColumns $plotMeasurementColumns `
                    -PlotAlias "pl" `
                    -PlotColumns $Columns
                if ([string]::IsNullOrWhiteSpace($statusExpression)) { continue }

                $condition = Get-MissingRequiredPlotFieldCondition -FieldExpression "pl.$(Quote-Name $fieldName)" -StatusExpression $statusExpression
                if ($canJoinPlotMeasurements) {
                    $periodScope = Get-SelectedPeriodAliasCondition -Alias "pm"
                    if (-not [string]::IsNullOrWhiteSpace($periodScope)) { $condition = "(($condition) AND ($periodScope))" }
                }
                $recordLabel = Get-RequiredPlotRecordLabelExpression `
                    -MeasurementAlias "" `
                    -MeasurementColumns @() `
                    -PlotAlias "pl" `
                    -PlotColumns $Columns `
                    -FallbackExpression "(pl.$(Quote-Name $targetIdFieldName) & '')"
                $observedExpression = if ($canJoinPlotMeasurements) {
                    Sql-Text "Selected-period PlotStatus 1 or 2; $displayName is blank"
                }
                else {
                    $observedSql.Replace("%STATUS%", $statusExpression)
                }
                $selectSql = @"
SELECT DISTINCT
    pl.$(Quote-Name $targetIdFieldName) AS [RowID],
    $recordLabel AS [RecordLabel],
    $observedExpression AS [ObservedValue]
INTO $(Quote-Name $tempName)
$fromSql
WHERE $condition
"@
            }
        }

        if ([string]::IsNullOrWhiteSpace($selectSql) -or [string]::IsNullOrWhiteSpace($targetIdFieldName)) { continue }

        $auditCount += Add-TempAuditFromSelect `
            -Connection $Connection `
            -TargetTableName $TableName `
            -TargetIdFieldName $targetIdFieldName `
            -TempTableName $tempName `
            -SelectIntoSql $selectSql `
            -RuleName "Required plot entry" `
            -FieldName $fieldName `
            -Message $message
    }

    return $auditCount
}

function Get-RequiredPlotEntryVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName
    )

    if ($TableName -notin @("Plots", "PlotMeasurements", "PlotCustomMeasurements")) { return "" }
    if ([string]::IsNullOrWhiteSpace($FieldName) -or $FieldName -match "/") { return "" }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName $TableName)) { return "" }

    $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    if (-not (Test-ColumnExists -Columns $columns -Name $FieldName)) { return "" }

    $plotColumns = @()
    $plotMeasurementColumns = @()
    try { if (Test-TableAvailable -Tables $tables -TableName "Plots") { $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots") } } catch { $plotColumns = @() }
    try { if (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") { $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements") } } catch { $plotMeasurementColumns = @() }

    switch ($TableName) {
        "PlotMeasurements" {
            $canJoinPlots = ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
                (Test-ColumnExists -Columns $columns -Name "PlotID") -and
                (Test-ColumnExists -Columns $plotColumns -Name "PlotID"))
            $fromSql = if ($canJoinPlots) {
                "FROM [PlotMeasurements] AS pm LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]"
            }
            else {
                "FROM [PlotMeasurements] AS pm"
            }
            $statusExpression = Get-RequiredPlotStatusExpression `
                -MeasurementAlias "pm" `
                -MeasurementColumns $columns `
                -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                -PlotColumns $plotColumns
            if ([string]::IsNullOrWhiteSpace($statusExpression)) { return "" }

            $whereClause = Get-MissingRequiredPlotFieldCondition -FieldExpression "pm.$(Quote-Name $FieldName)" -StatusExpression $statusExpression
            $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "pm"
            $selectList = Get-AliasSelectList -Connection $Connection -TableName "PlotMeasurements" -Alias "pm"
            $extraColumns = ", $statusExpression AS [ResolvedPlotStatus]"
            if ($canJoinPlots -and (Test-ColumnExists -Columns $plotColumns -Name "PlotNumber")) { $extraColumns += ", pl.[PlotNumber] AS [PlotNumber]" }
            return @"
SELECT $selectList$extraColumns
$fromSql
WHERE $whereClause;
"@
        }
        "PlotCustomMeasurements" {
            if (-not ((Test-ColumnExists -Columns $columns -Name "PlotMeasKey") -and (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotMeasKey"))) { return "" }
            $canJoinPlots = ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
                (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotID") -and
                (Test-ColumnExists -Columns $plotColumns -Name "PlotID"))
            $fromSql = if ($canJoinPlots) {
                "FROM ([PlotCustomMeasurements] AS c LEFT JOIN [PlotMeasurements] AS pm ON c.[PlotMeasKey] = pm.[PlotMeasKey]) LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]"
            }
            else {
                "FROM [PlotCustomMeasurements] AS c LEFT JOIN [PlotMeasurements] AS pm ON c.[PlotMeasKey] = pm.[PlotMeasKey]"
            }
            $statusExpression = Get-RequiredPlotStatusExpression `
                -MeasurementAlias "pm" `
                -MeasurementColumns $plotMeasurementColumns `
                -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                -PlotColumns $plotColumns
            if ([string]::IsNullOrWhiteSpace($statusExpression)) { return "" }

            $whereClause = Get-MissingRequiredPlotFieldCondition -FieldExpression "c.$(Quote-Name $FieldName)" -StatusExpression $statusExpression
            $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "pm"
            $selectList = Get-AliasSelectList -Connection $Connection -TableName "PlotCustomMeasurements" -Alias "c"
            $extraColumns = ", pm.[PeriodNumber] AS [PeriodNumber], $statusExpression AS [ResolvedPlotStatus]"
            if ($canJoinPlots -and (Test-ColumnExists -Columns $plotColumns -Name "PlotNumber")) { $extraColumns += ", pl.[PlotNumber] AS [PlotNumber]" }
            return @"
SELECT $selectList$extraColumns
$fromSql
WHERE $whereClause;
"@
        }
        "Plots" {
            $canJoinPlotMeasurements = ((Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") -and
                (Test-ColumnExists -Columns $columns -Name "PlotID") -and
                (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotID"))
            $fromSql = if ($canJoinPlotMeasurements) {
                "FROM [Plots] AS pl INNER JOIN [PlotMeasurements] AS pm ON pl.[PlotID] = pm.[PlotID]"
            }
            else {
                "FROM [Plots] AS pl"
            }
            $statusExpression = Get-RequiredPlotStatusExpression `
                -MeasurementAlias $(if ($canJoinPlotMeasurements) { "pm" } else { "" }) `
                -MeasurementColumns $plotMeasurementColumns `
                -PlotAlias "pl" `
                -PlotColumns $columns
            if ([string]::IsNullOrWhiteSpace($statusExpression)) { return "" }

            $whereClause = Get-MissingRequiredPlotFieldCondition -FieldExpression "pl.$(Quote-Name $FieldName)" -StatusExpression $statusExpression
            if ($canJoinPlotMeasurements) {
                $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "pm"
            }
            $selectList = Get-AliasSelectList -Connection $Connection -TableName "Plots" -Alias "pl"
            $extraColumns = ", $statusExpression AS [ResolvedPlotStatus]"
            if ($canJoinPlotMeasurements -and (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PeriodNumber")) { $extraColumns += ", pm.[PeriodNumber] AS [PeriodNumber]" }
            return @"
SELECT DISTINCT $selectList$extraColumns
$fromSql
WHERE $whereClause;
"@
        }
    }

    return ""
}

function Get-PlotStatusRemarkVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName
    )

    if ($TableName -notin @("Plots", "PlotMeasurements")) { return "" }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName $TableName)) { return "" }

    $columns = @(Get-TableColumns -Connection $Connection -TableName $TableName)
    $plotColumns = @()
    $plotMeasurementColumns = @()
    try { if (Test-TableAvailable -Tables $tables -TableName "Plots") { $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots") } } catch { $plotColumns = @() }
    try { if (Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") { $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements") } } catch { $plotMeasurementColumns = @() }
    $statusCandidates = @("PlotStatus", "PlotStatusCode", "PlotStatusID", "Status", "StatusCode", "StatusID")

    switch ($TableName) {
        "PlotMeasurements" {
            if ([string]::IsNullOrWhiteSpace((Get-FirstMatchingColumnName -Columns $columns -CandidateNames $statusCandidates))) { return "" }
            $canJoinPlots = ((Test-TableAvailable -Tables $tables -TableName "Plots") -and
                (Test-ColumnExists -Columns $columns -Name "PlotID") -and
                (Test-ColumnExists -Columns $plotColumns -Name "PlotID"))
            $fromSql = if ($canJoinPlots) {
                "FROM [PlotMeasurements] AS pm LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]"
            }
            else {
                "FROM [PlotMeasurements] AS pm"
            }
            $statusExpression = Get-RequiredPlotStatusExpression `
                -MeasurementAlias "pm" `
                -MeasurementColumns $columns `
                -PlotAlias $(if ($canJoinPlots) { "pl" } else { "" }) `
                -PlotColumns $plotColumns
            if ([string]::IsNullOrWhiteSpace($statusExpression)) { return "" }

            $remarksExpression = Get-RemarksValueExpression -Sources @(
                @{ Alias = "pm"; Columns = $columns },
                @{ Alias = $(if ($canJoinPlots) { "pl" } else { "" }); Columns = $plotColumns }
            )
            $whereClause = Get-MissingPlotStatusRemarkCondition -StatusExpression $statusExpression -RemarksExpression $remarksExpression
            $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "pm"
            $selectList = Get-AliasSelectList -Connection $Connection -TableName "PlotMeasurements" -Alias "pm"
            $extraColumns = ", $statusExpression AS [ResolvedPlotStatus], $remarksExpression AS [ResolvedPlotRemarks]"
            if ($canJoinPlots -and (Test-ColumnExists -Columns $plotColumns -Name "PlotNumber")) { $extraColumns += ", pl.[PlotNumber] AS [PlotNumber]" }
            return @"
SELECT $selectList$extraColumns
$fromSql
WHERE $whereClause;
"@
        }
        "Plots" {
            if (-not (Test-ColumnExists -Columns $columns -Name "PlotID")) { return "" }
            $canJoinPlotMeasurements = ((Test-TableAvailable -Tables $tables -TableName "PlotMeasurements") -and
                (Test-ColumnExists -Columns $columns -Name "PlotID") -and
                (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PlotID"))
            $fromSql = if ($canJoinPlotMeasurements) {
                "FROM [Plots] AS pl INNER JOIN [PlotMeasurements] AS pm ON pl.[PlotID] = pm.[PlotID]"
            }
            else {
                "FROM [Plots] AS pl"
            }
            $statusExpression = Get-RequiredPlotStatusExpression `
                -MeasurementAlias "" `
                -MeasurementColumns @() `
                -PlotAlias "pl" `
                -PlotColumns $columns
            if ([string]::IsNullOrWhiteSpace($statusExpression)) { return "" }

            $remarksExpression = Get-RemarksValueExpression -Sources @(
                @{ Alias = "pl"; Columns = $columns }
            )
            $whereClause = Get-MissingPlotStatusRemarkCondition -StatusExpression $statusExpression -RemarksExpression $remarksExpression
            if ($canJoinPlotMeasurements) {
                $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "pm"
            }
            $selectList = Get-AliasSelectList -Connection $Connection -TableName "Plots" -Alias "pl"
            $extraColumns = ", $statusExpression AS [ResolvedPlotStatus], $remarksExpression AS [ResolvedPlotRemarks]"
            if ($canJoinPlotMeasurements -and (Test-ColumnExists -Columns $plotMeasurementColumns -Name "PeriodNumber")) { $extraColumns += ", pm.[PeriodNumber] AS [PeriodNumber]" }
            return @"
SELECT DISTINCT $selectList$extraColumns
$fromSql
WHERE $whereClause;
"@
        }
    }

    return ""
}

function Add-RequiredTreeSpeciesChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    foreach ($requiredColumn in @("MeasurementID", "TreeID", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $Columns -Name $requiredColumn)) { return 0 }
    }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "Trees")) { return 0 }
    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "TreeID")) { return 0 }
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")) { return 0 }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "Trees" -FieldName "SpeciesCode")) { return 0 }

    $condition = Get-TreeSpeciesRequiredCondition -TreeHistoryField "tm.[TreeHistory]" -SpeciesField "treeMeta.[SpeciesCode]"
    $periodScope = Get-SelectedPeriodAliasCondition -Alias "tm"
    if (-not [string]::IsNullOrWhiteSpace($periodScope)) {
        $condition = "(($condition) AND ($periodScope))"
    }

    $recordLabelExpression = Get-TreeStemCountRecordLabelExpression -Columns $Columns -Alias "tm"
    $speciesTemp = "InventoryCleanTreeSpecies_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $speciesSelect = @"
SELECT
    tm.[MeasurementID] AS [RowID],
    $recordLabelExpression AS [RecordLabel],
    ('TreeHistory=' & (tm.[TreeHistory] & '') & '; SpeciesCode is blank') AS [ObservedValue]
INTO $(Quote-Name $speciesTemp)
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $condition
"@

    $message = Add-FieldManualTipToMessage `
        -Message "Trees.SpeciesCode is required for every tree measurement row in the selected cleaning period scope." `
        -TableName "Trees" `
        -FieldName "SpeciesCode"

    return Add-TempAuditFromSelect `
        -Connection $Connection `
        -TargetTableName "TreeMeasurements" `
        -TargetIdFieldName "MeasurementID" `
        -TempTableName $speciesTemp `
        -SelectIntoSql $speciesSelect `
        -RuleName "Required tree species" `
        -FieldName "SpeciesCode" `
        -Message $message
}

function Get-RequiredTreeSpeciesVerificationSql {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return "" }
    if (-not (Test-TableAvailable -Tables $tables -TableName "Trees")) { return "" }

    $measurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    foreach ($requiredColumn in @("TreeID", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $measurementColumns -Name $requiredColumn)) { return "" }
    }
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "TreeID")) { return "" }
    if (-not (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")) { return "" }

    $condition = Get-TreeSpeciesRequiredCondition -TreeHistoryField "tm.[TreeHistory]" -SpeciesField "treeMeta.[SpeciesCode]"
    $condition = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $condition -Alias "tm"
    $tmSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "tm"
    $extraColumns = ", treeMeta.[SpeciesCode] AS [TreeSpeciesCode]"
    if (Test-ColumnExists -Columns $treeColumns -Name "PlotNumber") { $extraColumns += ", treeMeta.[PlotNumber] AS [PlotNumber]" }
    if (Test-ColumnExists -Columns $treeColumns -Name "TreeNumber") { $extraColumns += ", treeMeta.[TreeNumber] AS [TreeNumber]" }

    return @"
SELECT $tmSelect$extraColumns
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $condition;
"@
}

function Get-RequiredRegenSpeciesVerificationSql {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "RegenMeasurements")) { return "" }

    $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements")
    if (-not (Test-ColumnExists -Columns $regenColumns -Name "SpeciesCode")) { return "" }

    $condition = Get-RegenSpeciesRequiredCondition -Columns $regenColumns -SpeciesField "[SpeciesCode]"
    if ([string]::IsNullOrWhiteSpace($condition)) { return "" }

    return New-ScopedSelectVerificationSql -Connection $Connection -TableName "RegenMeasurements" -RuleName "Required regen species" -FieldName "SpeciesCode" -Condition $condition
}

function Add-RequiredTreeCrownChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    foreach ($requiredColumn in @("MeasurementID", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $Columns -Name $requiredColumn)) { return 0 }
    }

    $recordLabel = Get-TreeStemCountRecordLabelExpression -Columns $Columns -Alias ""
    $auditCount = 0
    foreach ($definition in @(Get-RequiredTreeCrownFieldDefinitions)) {
        $fieldName = Get-FirstRequiredPlotFieldColumnName -Columns $Columns -Definition $definition
        if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }
        if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName $fieldName)) { continue }

        $field = Quote-Name $fieldName
        $displayName = [string]$definition.DisplayName
        $ruleName = [string]$definition.RuleName
        $condition = Get-TreeCrownRequiredCondition -TreeHistoryField "[TreeHistory]" -CrownField $field
        $message = Add-FieldManualTipToMessage `
            -Message "$displayName is required when TreeHistory is 0, 5, or 10. Valid nonblank values are checked against AppColumns/AppColumnCodes when those codes are configured." `
            -TableName "TreeMeasurements" `
            -FieldName $fieldName

        $auditCount += Add-ConditionAudit `
            -Connection $Connection `
            -Transaction $null `
            -TableName "TreeMeasurements" `
            -RuleName $ruleName `
            -FieldName $fieldName `
            -ObservedExpression "('TreeHistory=' & ([TreeHistory] & '') & '; $displayName is blank')" `
            -Message $message `
            -RecordLabelExpression $recordLabel `
            -Condition $condition
    }

    return $auditCount
}

function Get-RequiredTreeCrownVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$FieldName
    )

    if ([string]::IsNullOrWhiteSpace($FieldName) -or $FieldName -match "/") { return "" }
    if (-not (Test-RequiredTreeCrownFieldName -FieldName $FieldName)) { return "" }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return "" }
    $columns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredColumn in @("TreeHistory", $FieldName)) {
        if (-not (Test-ColumnExists -Columns $columns -Name $requiredColumn)) { return "" }
    }

    $condition = Get-TreeCrownRequiredCondition -TreeHistoryField "[TreeHistory]" -CrownField (Quote-Name $FieldName)
    return New-ScopedSelectVerificationSql -Connection $Connection -TableName "TreeMeasurements" -RuleName "Required crown value" -FieldName $FieldName -Condition $condition
}

function Add-RequiredTreeRadialIncrementChecks {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [object[]]$Columns
    )

    foreach ($requiredColumn in @("MeasurementID", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $Columns -Name $requiredColumn)) { return 0 }
    }

    $definition = @(Get-RequiredTreeRadialIncrementFieldDefinitions)[0]
    $fieldName = Get-FirstRequiredPlotFieldColumnName -Columns $Columns -Definition $definition
    if ([string]::IsNullOrWhiteSpace($fieldName)) { return 0 }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName "TreeMeasurements" -FieldName $fieldName)) { return 0 }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    $treeColumns = @()
    $canJoinTrees = $false
    if ((Test-TableAvailable -Tables $tables -TableName "Trees") -and
        (Test-ColumnExists -Columns $Columns -Name "TreeID")) {
        try {
            $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
            $canJoinTrees =
                (Test-ColumnExists -Columns $treeColumns -Name "TreeID") -and
                (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")
        }
        catch {
            $treeColumns = @()
            $canJoinTrees = $false
        }
    }

    $speciesField = if ($canJoinTrees) { "treeMeta.[SpeciesCode]" } else { "" }
    $problemFields = New-Object System.Collections.Generic.List[string]
    if (Test-ColumnExists -Columns $Columns -Name "Problem1") { [void]$problemFields.Add("tm.[Problem1]") }
    if (Test-ColumnExists -Columns $Columns -Name "Problem2") { [void]$problemFields.Add("tm.[Problem2]") }
    $condition = Get-TreeRadialIncrementRequiredCondition `
        -TreeHistoryField "tm.[TreeHistory]" `
        -RadialIncrementField "tm.$(Quote-Name $fieldName)" `
        -SpeciesField $speciesField `
        -ProblemFields ([string[]]$problemFields.ToArray())
    $condition = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $condition -Alias "tm"

    $recordLabelExpression = Get-TreeStemCountRecordLabelExpression -Columns $Columns -Alias "tm"
    $speciesObserved = if ($canJoinTrees) { "(treeMeta.[SpeciesCode] & '')" } else { "''" }
    $problem1Observed = if (Test-ColumnExists -Columns $Columns -Name "Problem1") { Get-ProblemCodeLabelExpression -ValueExpression "tm.[Problem1]" } else { "''" }
    $problem2Observed = if (Test-ColumnExists -Columns $Columns -Name "Problem2") { Get-ProblemCodeLabelExpression -ValueExpression "tm.[Problem2]" } else { "''" }
    $fromSql = if ($canJoinTrees) {
        @"
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
"@
    }
    else {
        "FROM [TreeMeasurements] AS tm"
    }

    $messageText = "Radial increment is required for timber trees with TreeHistory 5 (missed), and for timber trees with TreeHistory 0, 5, or 10 when Problem1 or Problem2 is 121 (negative diameter growth). Woodland species entered in the woodland species option are skipped."
    if (-not $canJoinTrees -and @(Get-WoodlandSpeciesCodeValues).Count -gt 0) {
        $messageText += " Woodland species could not be skipped because Trees.SpeciesCode could not be joined."
    }
    $message = Add-FieldManualTipToMessage `
        -Message $messageText `
        -TableName "TreeMeasurements" `
        -FieldName $fieldName

    $radialTemp = "InventoryCleanTreeRadial_" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
    $radialSelect = @"
SELECT
    tm.[MeasurementID] AS [RowID],
    $recordLabelExpression AS [RecordLabel],
    ('SpeciesCode=' & $speciesObserved & '; TreeHistory=' & (tm.[TreeHistory] & '') & '; Problem1=' & $problem1Observed & '; Problem2=' & $problem2Observed & '; $fieldName is blank') AS [ObservedValue]
INTO $(Quote-Name $radialTemp)
$fromSql
WHERE $condition
"@

    return Add-TempAuditFromSelect `
        -Connection $Connection `
        -TargetTableName "TreeMeasurements" `
        -TargetIdFieldName "MeasurementID" `
        -TempTableName $radialTemp `
        -SelectIntoSql $radialSelect `
        -RuleName "Required radial increment" `
        -FieldName $fieldName `
        -Message $message
}

function Get-RequiredTreeRadialIncrementVerificationSql {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$FieldName
    )

    if ([string]::IsNullOrWhiteSpace($FieldName) -or $FieldName -match "/") { return "" }
    if (-not (Test-RequiredTreeRadialIncrementFieldName -FieldName $FieldName)) { return "" }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not (Test-TableAvailable -Tables $tables -TableName "TreeMeasurements")) { return "" }
    $columns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    foreach ($requiredColumn in @("TreeHistory", $FieldName)) {
        if (-not (Test-ColumnExists -Columns $columns -Name $requiredColumn)) { return "" }
    }

    $treeColumns = @()
    $canJoinTrees = $false
    if ((Test-TableAvailable -Tables $tables -TableName "Trees") -and
        (Test-ColumnExists -Columns $columns -Name "TreeID")) {
        try {
            $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
            $canJoinTrees =
                (Test-ColumnExists -Columns $treeColumns -Name "TreeID") -and
                (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")
        }
        catch {
            $treeColumns = @()
            $canJoinTrees = $false
        }
    }

    $speciesField = if ($canJoinTrees) { "treeMeta.[SpeciesCode]" } else { "" }
    $problemFields = New-Object System.Collections.Generic.List[string]
    if (Test-ColumnExists -Columns $columns -Name "Problem1") { [void]$problemFields.Add("tm.[Problem1]") }
    if (Test-ColumnExists -Columns $columns -Name "Problem2") { [void]$problemFields.Add("tm.[Problem2]") }
    $condition = Get-TreeRadialIncrementRequiredCondition `
        -TreeHistoryField "tm.[TreeHistory]" `
        -RadialIncrementField "tm.$(Quote-Name $FieldName)" `
        -SpeciesField $speciesField `
        -ProblemFields ([string[]]$problemFields.ToArray())
    $condition = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $condition -Alias "tm"
    $tmSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "tm"

    if ($canJoinTrees) {
        $extraColumns = ", treeMeta.[SpeciesCode] AS [TreeSpeciesCode]"
        if (Test-ColumnExists -Columns $treeColumns -Name "PlotNumber") { $extraColumns += ", treeMeta.[PlotNumber] AS [PlotNumber]" }
        if (Test-ColumnExists -Columns $treeColumns -Name "TreeNumber") { $extraColumns += ", treeMeta.[TreeNumber] AS [TreeNumber]" }
        return @"
SELECT $tmSelect$extraColumns
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $condition;
"@
    }

    return @"
SELECT $tmSelect
FROM [TreeMeasurements] AS tm
WHERE $condition;
"@
}

function Add-CfiRangeChecksForTable {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [object[]]$Columns
    )

    $plotField = Find-CandidateColumn -Columns $Columns -Patterns @("^plotmeaskey$", "^plotkey$", "^plotid$", "^plotnumber$", "^plot$", "plotid", "plotnumber", "stand", "unit")
    $treeField = Find-CandidateColumn -Columns $Columns -Patterns @("^treemeaskey$", "^regenmeaskey$", "^treekey$", "^treeid$", "^treenumber$", "^tree$", "tag", "stem")
    $recordLabel = Get-RecordLabelExpression -PlotField $plotField -TreeField $treeField
    $auditCount = 0
    $idbhMax = Get-ControlDecimalValue -Control $dbhMax -DefaultValue 500
    $heightMaximum = Get-ControlDecimalValue -Control $heightMax -DefaultValue 150
    $stemCountMaximum = Get-StemCountMaximumValue
    $idbhMaxSql = $idbhMax.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $heightMaxSql = $heightMaximum.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $stemCountMaxSql = $stemCountMaximum.ToString([System.Globalization.CultureInfo]::InvariantCulture)

    if ($TableName -eq "PlotMeasurements") {
        $auditCount += Add-PlotRuleCoverageDiagnostics -Connection $Connection
    }

    if ($TableName -eq "TreeMeasurements") {
        $auditCount += Add-TreeRuleCoverageDiagnostics -Connection $Connection
    }

    if ($TableName -eq "RegenMeasurements") {
        $auditCount += Add-RegenRuleCoverageDiagnostics -Connection $Connection -Columns $Columns
        $auditCount += Add-RequiredRegenSpeciesChecks -Connection $Connection -Columns $Columns
    }

    $shouldCheckIdbh = (Test-ColumnExists -Columns $Columns -Name "IDBH") -and
        (($TableName -eq "RegenMeasurements") -or
            (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName "IDBH"))
    if ($shouldCheckIdbh) {
        if ($TableName -eq "RegenMeasurements") {
            if (Test-ColumnExists -Columns $Columns -Name "SpeciesCode") {
                $regenIdbhMessage = Add-FieldManualTipToMessage `
                    -Message "When RegenMeasurements.SpeciesCode is greater than 0, Regen IDBH is required and must be one of the valid class codes: 0, 20, or 40." `
                    -TableName $TableName `
                    -FieldName "IDBH"

                $auditCount += Add-ConditionAudit `
                    -Connection $Connection `
                    -Transaction $null `
                    -TableName $TableName `
                    -RuleName "Regen IDBH code check" `
                    -FieldName "IDBH" `
                    -ObservedExpression "('SpeciesCode=' & ([SpeciesCode] & '') & '; IDBH=' & ([IDBH] & ''))" `
                    -Message $regenIdbhMessage `
                    -RecordLabelExpression $recordLabel `
                    -Condition (Get-RegenIdbhRequiredCondition -SpeciesField "[SpeciesCode]" -IdbhField "[IDBH]")
            }
        }
        else {
            $idbhRangeMessage = Add-FieldManualTipToMessage `
                -Message "Entered IDBH values must be greater than zero and no more than the configured maximum." `
                -TableName $TableName `
                -FieldName "IDBH"

            if ($TableName -eq "TreeMeasurements" -and (Test-ColumnExists -Columns $Columns -Name "TreeHistory")) {
                $missingRequiredCondition = Get-TreeIdbhMissingRequiredCondition -TreeHistoryField "[TreeHistory]" -IdbhField "[IDBH]"
                $shouldBeBlankCondition = Get-TreeIdbhShouldBeBlankCondition -TreeHistoryField "[TreeHistory]" -IdbhField "[IDBH]"
                $missingDbhMessage = Add-FieldManualTipToMessage `
                    -Message "TreeMeasurements.IDBH is blank, but this TreeHistory status is selected in Run options as requiring DBH." `
                    -TableName $TableName `
                    -FieldName "IDBH"
                $dbhShouldBeBlankMessage = Add-FieldManualTipToMessage `
                    -Message "TreeMeasurements.IDBH should be blank for TreeHistory 1, 4, or 8." `
                    -TableName $TableName `
                    -FieldName "IDBH"
                $auditCount += Add-ConditionAudit `
                    -Connection $Connection `
                    -Transaction $null `
                    -TableName $TableName `
                    -RuleName "Missing DBH" `
                    -FieldName "IDBH" `
                    -ObservedExpression "[IDBH]" `
                    -Message $missingDbhMessage `
                    -RecordLabelExpression $recordLabel `
                    -Condition $missingRequiredCondition

                $auditCount += Add-ConditionAudit `
                    -Connection $Connection `
                    -Transaction $null `
                    -TableName $TableName `
                    -RuleName "DBH should be blank" `
                    -FieldName "IDBH" `
                    -ObservedExpression "('TreeHistory=' & ([TreeHistory] & '') & '; IDBH=' & ([IDBH] & ''))" `
                    -Message $dbhShouldBeBlankMessage `
                    -RecordLabelExpression $recordLabel `
                    -Condition $shouldBeBlankCondition

                $auditCount += Add-ConditionAudit `
                    -Connection $Connection `
                    -Transaction $null `
                    -TableName $TableName `
                    -RuleName "IDBH range check" `
                    -FieldName "IDBH" `
                    -ObservedExpression "[IDBH]" `
                    -Message $idbhRangeMessage `
                    -RecordLabelExpression $recordLabel `
                    -Condition "[IDBH] Is Not Null AND ([IDBH] <= 0 OR [IDBH] > $idbhMaxSql)"
            }
            else {
                $auditCount += Add-RangeAudit `
                    -Connection $Connection `
                    -Transaction $null `
                    -TableName $TableName `
                    -FieldName "IDBH" `
                    -MaximumValue $idbhMax `
                    -RuleName "IDBH range check" `
                    -Message $idbhRangeMessage `
                    -RecordLabelExpression $recordLabel
            }
        }
    }

    if ($TableName -eq "RegenMeasurements") {
        $auditCount += Add-RegenMinorPlotChecks -Connection $Connection -Columns $Columns
    }

    if ((Test-ColumnExists -Columns $Columns -Name "TotalHeight") -and
        (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName "TotalHeight")) {
        if ($TableName -eq "TreeMeasurements") {
            $auditCount += Add-TreeHeightChecks -Connection $Connection -Columns $Columns
            $auditCount += Add-TreeHeightSubsampleProtocolChecks -Connection $Connection -Columns $Columns
        }
        else {
            $heightMessage = Add-FieldManualTipToMessage `
                -Message "TotalHeight is zero, negative, or above the configured maximum when present." `
                -TableName $TableName `
                -FieldName "TotalHeight"

            $auditCount += Add-ConditionAudit `
                -Connection $Connection `
                -Transaction $null `
                -TableName $TableName `
                -RuleName "Height range check" `
                -FieldName "TotalHeight" `
                -ObservedExpression "[TotalHeight]" `
                -Message $heightMessage `
                -RecordLabelExpression $recordLabel `
                -Condition "[TotalHeight] Is Not Null AND ([TotalHeight] <= 0 OR [TotalHeight] > $heightMaxSql)"
        }
    }

    $shouldCheckStemCount = (Test-ColumnExists -Columns $Columns -Name "StemCount") -and
        (($TableName -eq "RegenMeasurements") -or
            (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName "StemCount"))
    if ($shouldCheckStemCount) {
        if ($TableName -eq "RegenMeasurements") {
            if (-not (Test-ColumnExists -Columns $Columns -Name "SpeciesCode")) { return $auditCount }
            $stemMessage = Add-FieldManualTipToMessage `
                -Message "Entered Regen StemCount values must be greater than 0 and no more than the configured maximum. When RegenMeasurements.SpeciesCode is greater than 0, Regen StemCount is required." `
                -TableName $TableName `
                -FieldName "StemCount"

            $auditCount += Add-ConditionAudit `
                -Connection $Connection `
                -Transaction $null `
                -TableName $TableName `
                -RuleName "Stem count range check" `
                -FieldName "StemCount" `
                -ObservedExpression "('SpeciesCode=' & ([SpeciesCode] & '') & '; StemCount=' & ([StemCount] & ''))" `
                -Message $stemMessage `
                -RecordLabelExpression $recordLabel `
                -Condition (Get-RegenStemCountRequiredCondition -SpeciesField "[SpeciesCode]" -StemCountField "[StemCount]" -MaximumValue $stemCountMaximum)
        }
        elseif ($TableName -eq "TreeMeasurements" -and (Test-TreeStemCountChecksEnabled)) {
            $auditCount += Add-TreeStemCountChecks -Connection $Connection -Columns $Columns
        }
    }

    if ($TableName -in @("Plots", "PlotMeasurements", "PlotCustomMeasurements")) {
        $auditCount += Add-RequiredPlotStatusFieldChecks -Connection $Connection -TableName $TableName -Columns $Columns
        $auditCount += Add-PlotStatusRemarkChecks -Connection $Connection -TableName $TableName -Columns $Columns
    }

    if ($TableName -eq "TreeMeasurements") {
        $auditCount += Add-RequiredTreeSpeciesChecks -Connection $Connection -Columns $Columns
        $auditCount += Add-RequiredTreeCrownChecks -Connection $Connection -Columns $Columns
        $auditCount += Add-RequiredTreeRadialIncrementChecks -Connection $Connection -Columns $Columns
    }

    return $auditCount
}

function Get-ValidCodesSqlList {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName
    )

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not ($tables -contains "AppTables") -or
        -not ($tables -contains "AppColumns") -or
        -not ($tables -contains "AppColumnCodes")) {
        return ""
    }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName $FieldName)) {
        return ""
    }

    $metadataSql = @"
SELECT c.[ID]
FROM [AppTables] AS t
INNER JOIN [AppColumns] AS c ON t.[ID] = c.[TableID]
WHERE t.[TableName] = $(Sql-Text $TableName)
  AND c.[ColumnName] = $(Sql-Text $FieldName)
  AND c.[CategoryTypeID] = 3
"@
    $metadata = Get-DataTable -Connection $Connection -Sql $metadataSql
    if ($metadata.Rows.Count -eq 0) { return "" }

    $validCodes = New-Object System.Collections.Generic.List[string]
    foreach ($row in $metadata.Rows) {
        $columnId = [string]$row["ID"]
        $codes = Get-DataTable -Connection $Connection -Sql "SELECT [Code] FROM [AppColumnCodes] WHERE [ColumnID] = $(Sql-Text $columnId)"
        foreach ($codeRow in $codes.Rows) {
            $code = Normalize-CodeValue $codeRow["Code"]
            Add-UniqueCode -Codes $validCodes -Code $code
            if ($code -match "^\d+$") {
                Add-UniqueCode -Codes $validCodes -Code (([int64]$code).ToString())
            }
        }
    }

    if ($validCodes.Count -eq 0) { return "" }
    return (($validCodes | Sort-Object -Unique | ForEach-Object { Sql-Text $_ }) -join ", ")
}

function Get-CodedCleaningRulesSummary {
    $idbhMax = Get-ControlDecimalValue -Control $dbhMax -DefaultValue 500
    $heightMaximum = Get-ControlDecimalValue -Control $heightMax -DefaultValue 150
    $stemCountMaximum = Get-StemCountMaximumValue
    $idbhJump = Get-ControlDecimalValue -Control $dbhGrowthMax -DefaultValue 100
    $heightJump = Get-ControlDecimalValue -Control $heightGrowthMax -DefaultValue 20
    $woodlandCodes = @(Get-WoodlandSpeciesCodeValues)
    $woodlandText = if ($woodlandCodes.Count -gt 0) { [string]::Join(", ", [string[]]$woodlandCodes) } else { "none entered" }
    $treeStemCheckText = if (Test-TreeStemCountChecksEnabled) { "on" } else { "off" }
    $newMortalityIdbhText = if (Test-NewMortalityIdbhChecksEnabled) { "on" } else { "off" }
    $oldMortalityIdbhText = if (Test-OldMortalityIdbhChecksEnabled) { "on" } else { "off" }
    $newMortalityHeightText = if (Test-NewMortalityHeightChecksEnabled) { "on" } else { "off" }
    $oldMortalityHeightText = if (Test-OldMortalityHeightChecksEnabled) { "on" } else { "off" }
    $problemHeightCodesText = Get-ProblemHeightNoHeightCodesText
    $requiredTreeIdbhText = Get-TreeIdbhRequiredHistoryText
    $requiredTreeHeightText = Get-TreeHeightRequiredHistoryText
    $heightProtocolText = Get-TotalHeightProtocolDisplayText
    $rareSpeciesHeightText = Get-HeightRareSpeciesDisplayText
    $heightSubsampleText = "minimum count $(Get-HeightSubsampleMinimumCount), minimum eligible IDBH $(Get-HeightSubsampleMinimumIdbh), all heights at/above IDBH $(if (Test-HeightSubsampleAllAtOrAboveEnabled) { Get-HeightSubsampleAllAtOrAboveIdbh } else { 'off' })"
    $heightMinorPlots = @(Get-HeightRequiredMinorPlotValues)
    $heightMinorPlotText = if ($heightMinorPlots.Count -gt 0) { [string]::Join(", ", [string[]]$heightMinorPlots) } else { "none entered; all minor plots are included when height is required" }
    $regenTimberSeedlingPlots = @(Get-RegenTimberSeedlingMinorPlotValues)
    $regenTimberSapling20Plots = @(Get-RegenTimberSapling20MinorPlotValues)
    $regenTimberSapling40Plots = @(Get-RegenTimberSapling40MinorPlotValues)
    $regenWoodlandSeedlingPlots = @(Get-RegenWoodlandSeedlingMinorPlotValues)
    $regenWoodlandSapling20Plots = @(Get-RegenWoodlandSapling20MinorPlotValues)
    $regenWoodlandSapling40Plots = @(Get-RegenWoodlandSapling40MinorPlotValues)

    $lines = @(
        "Deterministic cleaning rules coded into this app:",
        "- " + (Get-CleaningPeriodScopeDescription),
        "- AppColumns active-field scope: only fields marked Active in AppColumns are reviewed for data-entry cleaning. Excluded fields such as per-acre expansion, CalcSiteIndex, and GMP are still skipped even if active.",
        "- CFI code checks: nonblank coded fields must match the template AppColumns/AppColumnCodes values.",
        "- Required plot entries: active plot fields for elevation, aspect, slope percent, slope position, UTM northing/easting/zone, measurement date, crew, stand class, stand age, stockability percent, and stockability factor are checked for missing values when PlotStatus is 1 or 2. Plot remarks are required when PlotStatus is 7.",
        "- PlotStatus progression: the first nonblank PlotStatus for a plot should be 2 (initial install); before that only blank/null is allowed. After PlotStatus 2, later periods must have PlotStatus 1, 4, 5, 6, or 7 and cannot be blank. PlotStatus progression is checked across all periods.",
        "- Not-measured/dropped plot data: active PlotMeasurements and PlotCustomMeasurements fields are flagged when data is recorded for PlotStatus 4 (missing plot), 5 (not measured), 6 (dropped - off reservation), or 7 (dropped - other). Plot remarks, UTM coordinates, management unit, and FLCCommercial are allowed.",
        "- Not-measured/dropped plot regen data: active RegenMeasurements and RegenCustomMeasurements fields are flagged when regen data is recorded for the matching plot period with PlotStatus 4 (missing plot), 5 (not measured), 6 (dropped - off reservation), or 7 (dropped - other). Regen remarks are allowed.",
        "- Per-acre expansion, CalcSiteIndex, and GMP fields are not treated as cleaning items and are skipped by code validation.",
        "- DMUML code check: DMUML is treated as a three-position numeric code. Digits 0, 1, 2, and 3 are allowed, and the three digits must add up to no more than 6. Short numeric entries are left-padded before checking, so 20 is read as 020.",
        "- Regen species: RegenMeasurements.SpeciesCode is required when a regen row has RegenMeasKey, IDBH, StemCount, or MinorPlot data recorded.",
        "- Regen IDBH: when RegenMeasurements.SpeciesCode is greater than 0, RegenMeasurements.IDBH is required and must be one of 0, 20, or 40; 0 is a valid regen DBH class code.",
        "- Regen MinorPlot: when RegenMeasurements.SpeciesCode is greater than 0, MinorPlot is required. Optional project rules can also restrict timber seedlings (IDBH 0), timber saplings IDBH 20, timber saplings IDBH 40, woodland seedlings (IDBH 0), woodland saplings IDBH 20, and woodland saplings IDBH 40 to entered minor plots. Current allowed minor plots: timber seedlings=$([string]::Join(', ', [string[]]$regenTimberSeedlingPlots)); timber saplings IDBH 20=$([string]::Join(', ', [string[]]$regenTimberSapling20Plots)); timber saplings IDBH 40=$([string]::Join(', ', [string[]]$regenTimberSapling40Plots)); woodland seedlings=$([string]::Join(', ', [string[]]$regenWoodlandSeedlingPlots)); woodland saplings IDBH 20=$([string]::Join(', ', [string[]]$regenWoodlandSapling20Plots)); woodland saplings IDBH 40=$([string]::Join(', ', [string[]]$regenWoodlandSapling40Plots)). Blank categories are not enforced for allowed-value checks. Woodland categories use the woodland species codes option: $woodlandText.",
        "- Tree species: Trees.SpeciesCode is required for every tree measurement row in the selected cleaning period scope.",
        "- Tree IDBH range: entered TreeMeasurements.IDBH values must be greater than 0 and no more than $idbhMax. Blank IDBH is flagged for $requiredTreeIdbhText. TreeHistory 1, 4, and 8 should not have DBH recorded; TreeHistory 6 and 9 may have blank IDBH. New mortality DBH check for TreeHistory 2/3 is $newMortalityIdbhText; old mortality DBH check for TreeHistory 7 is $oldMortalityIdbhText.",
        "- Tree crown fields: crown ratio and crown class are required for TreeHistory 0, 5, and 10 when those fields exist and are active in AppColumns. Nonblank invalid codes are checked by the template code-validation rules when AppColumnCodes are configured.",
        "- Tree radial increment: when the radial increment field exists and is active in AppColumns, radial increment is required for timber trees with TreeHistory 5 (missed), and for timber trees with TreeHistory 0, 5, or 10 when Problem1 or Problem2 is 121 (negative diameter growth). Woodland species entered in the woodland species option are skipped.",
        "- Height range and required-height protocol: entered TreeMeasurements.TotalHeight values must be greater than 0 and no more than $heightMaximum. Total Height Protocol is $heightProtocolText. Blank TotalHeight is flagged for $requiredTreeHeightText, limited to these minor plots when entered: $heightMinorPlotText. Rare species requiring 100% live-tree height: $rareSpeciesHeightText. Subsample settings: $heightSubsampleText. On older not-yet-crosswalked data, TreeClass 1/2/3/4/9 and TreeStatus 1/2/3 are treated as live-tree status for TotalHeight checks. Blank/null TreeHistory and TreeHistory 1, 4, 6, 8, and 9 may have blank height. New mortality height check for TreeHistory 2/3 is $newMortalityHeightText; old mortality height check for TreeHistory 7 is $oldMortalityHeightText. Selected no-height problem-code exceptions include newer 123/127/128 and legacy 72/74/75 when their matching option is on.",
        "- Not-measured/dropped plot tree data: active TreeMeasurements and TreeCustomMeasurements fields are flagged when data is recorded for the matching tree period on a plot with PlotStatus 4 (missing plot), 5 (not measured), 6 (dropped - off reservation), or 7 (dropped - other).",
        "- Problem-code height rule: for timber trees, entered TotalHeight is flagged when Problem1 or Problem2 is one of these selected no-height codes: $problemHeightCodesText. Blank height is treated as correct for these problem-code cases, including TreeHistory 10 ingrowth. Woodland species entered in the woodland species option are skipped.",
        "- New mortality Severity1 rule: when TreeHistory is 2 or 3 and Problem1 is recorded, Severity1 must be code 3.",
        "- Regen stem count range: entered RegenMeasurements.StemCount values must be greater than 0 and no more than $stemCountMaximum, even when SpeciesCode is blank. When SpeciesCode is greater than 0, StemCount must also be present.",
        "- Optional tree StemCount check: currently $treeStemCheckText. When on, entered TreeMeasurements.StemCount values must be 1 to $stemCountMaximum. Timber species may have blank StemCount; woodland/non-timber species entered in the woodland species box require StemCount. Current woodland species option: $woodlandText.",
        "- Remeasurement checks: IDBH increases greater than $idbhJump and TotalHeight increases greater than $heightJump are flagged for review. Shorter current TotalHeight is allowed when a previously live tree becomes a mortality TreeHistory status selected in Run options as requiring height, because dead trees can lose height through breakage.",
        "- Shrinking IDBH: only live timber trees are checked; woodland species codes are allowed to shrink and are skipped. Current woodland species exclusion option: $woodlandText.",
        "- TreeHistory checks: invalid live/mortality/harvest/thin/old mortality transitions and Lazarus-style returns to live/missed/ingrowth are flagged. TreeHistory 9 include/non-include corrections between periods are allowed.",
        "- Ingrowth checks: TreeHistory 10 is checked against plot status and measurement timing. TreeHistory 10 on install plots with PlotStatus 2 is flagged; TreeHistory 10 in the first project measurement period is flagged; TreeHistory 10 with an earlier nonblank TreeHistory value is flagged for review; earlier blank TreeHistory rows do not count as prior history for that rule; first recorded live TreeHistory 0 rows on remeasurement plots with PlotStatus 1 are flagged as possible missed ingrowth.",
        "- TreeClass/TreeStatus conversion checks: when TreeHistory and a detected TreeClass or TreeStatus field are active in AppColumns, the source code is compared with TreeHistory using the supplied crosswalk. Clear mismatches are flagged, and conditional crosswalk cases are exported as review rows instead of guessed corrections.",
        "- Problem/severity checks: Problem1/Severity1 and Problem2/Severity2 must agree. Blank, Null, and 0 mean none; values greater than 0 mean entered.",
        "- Missing measurement checks: trees and plots with measurements are checked for missing expected period rows.",
        "- Workbook count checks: plot, tree, regen, and custom measurement row counts are checked against the cleaning workbook rules."
    )

    return [string]::Join("`r`n", [string[]]$lines)
}

function Get-SpecificCodedRuleContext {
    param(
        [string]$TableName,
        [string]$FieldName,
        [string]$RuleName
    )

    $idbhMax = Get-ControlDecimalValue -Control $dbhMax -DefaultValue 500
    $heightMaximum = Get-ControlDecimalValue -Control $heightMax -DefaultValue 150
    $stemCountMaximum = Get-StemCountMaximumValue
    $idbhJump = Get-ControlDecimalValue -Control $dbhGrowthMax -DefaultValue 100
    $heightJump = Get-ControlDecimalValue -Control $heightGrowthMax -DefaultValue 20
    $woodlandCodes = @(Get-WoodlandSpeciesCodeValues)
    $woodlandText = if ($woodlandCodes.Count -gt 0) { [string]::Join(", ", [string[]]$woodlandCodes) } else { "none entered" }

    switch ($RuleName) {
        "DMUML code check" {
            return "This finding means $TableName.$FieldName is not a valid three-position DMUML code. Valid examples include 000, 001, 020, 200, 300, and combinations where each digit is 0, 1, 2, or 3 and the digit total is no more than 6."
        }
        "Invalid CFI code" {
            return "This finding means $TableName.$FieldName contains a nonblank value that is not listed as a valid code in the project template metadata."
        }
        "Regen IDBH code check" {
            return "This finding is regen-specific. RegenMeasurements.SpeciesCode is greater than 0, so RegenMeasurements.IDBH is required. IDBH is not a timber diameter; it is a class code and only 0, 20, and 40 are valid. Do not flag 0 as an error."
        }
        "Required regen minor plot" {
            return "This finding means RegenMeasurements.MinorPlot is blank even though SpeciesCode is greater than 0. DADA treats MinorPlot as required for active regen records. If the project also has specific seedling/sapling minor-plot rules entered, nonblank values outside those allowed lists are reported separately."
        }
        "Regen minor plot check" {
            return "This finding means RegenMeasurements.MinorPlot is nonblank but not one of the project-entered minor plots for that regen class. IDBH 0 is treated as seedlings; IDBH 20 and IDBH 40 are checked as separate sapling classes. Woodland/timber classification uses the woodland species code list entered in Run options."
        }
        "Rule coverage diagnostic" {
            return "This finding is a setup/coverage diagnostic, not a direct value correction. DADA found plot, tree, or regen data that may not be fully reviewable because a needed field, AppColumns active flag, join key, woodland species list, or project-specific allowed list is missing. Review the RecordedValue and message, update the run setup or AppColumns if needed, and rerun."
        }
        "Required regen species" {
            return "This finding means RegenMeasurements.SpeciesCode is blank even though the regen row has other regen data recorded, such as RegenMeasKey, IDBH, StemCount, or MinorPlot. Enter the correct species code from the source tally, or clear the row if no regen should be represented."
        }
        "Required plot entry" {
            return "This finding means $TableName.$FieldName is blank even though the plot has PlotStatus code 1 or 2. The app treats this as a required plot-entry field when it is present and active in AppColumns."
        }
        "Required plot remark" {
            return "This finding means the plot has PlotStatus code 7 but no plot remark was recorded. Add the plot remark explaining the status, or correct PlotStatus if code 7 was entered by mistake."
        }
        "PlotStatus progression check" {
            return "This finding means the PlotStatus sequence does not follow the expected progression. Periods before install should be blank, the first nonblank PlotStatus should be 2, and periods after install should be 1, 4, 5, 6, or 7 with no blanks."
        }
        "Plot data on not-measured plot" {
            return "This finding means $TableName.$FieldName has a recorded value even though the plot has PlotStatus 4, 5, 6, or 7. Those statuses mean missing, not measured, dropped off reservation, or dropped other. Plot remarks, UTM coordinates, management unit, and FLCCommercial are allowed because they may be preloaded or needed for dropped-plot documentation."
        }
        "Regen data on not-measured plot" {
            return "This finding means $TableName.$FieldName has a recorded value even though the matching plot measurement has PlotStatus 4, 5, 6, or 7. Those statuses mean missing, not measured, dropped off reservation, or dropped other, so regen measurement/custom data should be blank for that plot period. The PlotStatus column is included in the export for verification."
        }
        "Missing DBH" {
            return "This finding applies to TreeMeasurements.IDBH. DBH is blank, but this TreeHistory status is selected in Run options as requiring DBH. TreeHistory 1, 4, and 8 should have blank DBH; TreeHistory 6 and 9 may have blank DBH."
        }
        "DBH should be blank" {
            return "This finding applies to TreeMeasurements.IDBH. TreeHistory 1, 4, and 8 should not have DBH recorded, so the DBH value should be reviewed and removed unless TreeHistory is wrong."
        }
        "Required tree species" {
            return "This finding means the joined Trees.SpeciesCode value is blank for a tree measurement row in the selected cleaning period scope. Every measured tree needs a tree species for verification and reporting, even if TreeHistory is blank."
        }
        "Required crown ratio" {
            return "This finding means TreeMeasurements.$FieldName is blank for TreeHistory 0, 5, or 10. Crown ratio is required for those TreeHistory values. If a nonblank crown ratio is invalid, the template code-validation rule should also flag it when AppColumnCodes are configured."
        }
        "Required crown class" {
            return "This finding means TreeMeasurements.$FieldName is blank for TreeHistory 0, 5, or 10. Crown class is required for those TreeHistory values. If a nonblank crown class is invalid, the template code-validation rule should also flag it when AppColumnCodes are configured."
        }
        "Required radial increment" {
            return "This finding means TreeMeasurements.$FieldName is blank for a timber tree that requires radial increment: TreeHistory 5 (missed), or TreeHistory 0, 5, or 10 with Problem1/Problem2 code 121 (negative diameter growth). Woodland species entered in the woodland species option are skipped when Trees.SpeciesCode can be joined."
        }
        "Tree data on not-measured plot" {
            return "This finding means $TableName.$FieldName has a recorded value even though the matching plot measurement has PlotStatus 4, 5, 6, or 7. Those statuses mean missing, not measured, dropped off reservation, or dropped other, so tree measurement/custom data should be blank for that plot period."
        }
        "IDBH range check" {
            if ($TableName -eq "TreeMeasurements") {
                return "This finding applies to TreeMeasurements.IDBH. Entered DBH must be greater than 0 and no more than $idbhMax. Blank required DBH is reported separately as Missing DBH."
            }
            return "This finding means $TableName.$FieldName is missing, zero, negative, or greater than the configured IDBH maximum of $idbhMax."
        }
        "Height range check" {
            if ($TableName -eq "TreeMeasurements") {
            return "This finding applies to TreeMeasurements.TotalHeight. Entered height must be greater than 0 and no more than $heightMaximum. Total Height Protocol is " + (Get-TotalHeightProtocolDisplayText) + ". Blank height is flagged for " + (Get-TreeHeightRequiredHistoryText) + " and for rare species entered as 100% live-tree height species (" + (Get-HeightRareSpeciesDisplayText) + "). In subsample mode, ordinary live-tree blanks are reviewed by grouped plot/species/2-inch IDBH class counts instead of row-by-row missing-height errors. On older not-yet-crosswalked data, TreeClass 1/2/3/4/9 and TreeStatus 1/2/3 are treated as live-tree status for TotalHeight checks. Blank/null TreeHistory and TreeHistory 1, 4, 6, 8, and 9 may have blank height. Timber trees with entered TotalHeight and selected Problem1/Problem2 no-height codes (" + (Get-ProblemHeightNoHeightCodesText) + ") are also flagged; blank height is correct for those problem-code cases, including TreeHistory 10 ingrowth. Woodland species entered in the woodland species option are skipped. Minor plot filter: " + $(if ((Get-HeightRequiredMinorPlotValues).Count -gt 0) { [string]::Join(", ", [string[]](Get-HeightRequiredMinorPlotValues)) } else { "none; all minor plots are included when height is required" })
            }
            return "This finding means $TableName.$FieldName is present but zero, negative, or greater than the configured TotalHeight maximum of $heightMaximum."
        }
        "TotalHeight subsample review" {
            return "This finding is a grouped TotalHeight subsample shortage. DADA grouped eligible live trees by plot, SpeciesCode, and 2-inch IDBH class, then compared recorded height counts against the subsample minimum. TreeClass 1/2/3/4/9 and TreeStatus 1/2/3 can stand in as live status when old data has not been crosswalked. Rare species entered as 100% height species and selected problem-code no-height exception trees are excluded from this grouped count."
        }
        "TotalHeight subsample order review" {
            return "This finding means the plot/species/2-inch IDBH class had enough recorded heights overall, but one of the first required trees by TreeNumber had blank TotalHeight while a later TreeNumber was measured. DADA treats TreeNumber as starting at north and increasing clockwise."
        }
        "Stem count range check" {
            if ($TableName -eq "RegenMeasurements") {
                return "This finding means RegenMeasurements.StemCount is entered outside the valid range of 1 to $stemCountMaximum, or SpeciesCode is greater than 0 and StemCount is missing. Entered stem counts are checked even when SpeciesCode is blank, so a value such as 101 is still flagged."
            }
            return "This finding means $TableName.$FieldName is missing, zero, negative, or greater than the configured maximum regen stem count of $stemCountMaximum."
        }
        "Tree stem count check" {
            return "This finding applies to TreeMeasurements.StemCount. Entered tree stem counts must be 1 to $stemCountMaximum. Timber species may have blank StemCount; woodland/non-timber species entered in the woodland species box require StemCount. Current woodland species option: $woodlandText."
        }
        "IDBH jump check" {
            return "This finding compares the current tree measurement with the prior period and flags IDBH growth over $idbhJump. Shrinking IDBH on live timber trees is reported separately by the selected-period IDBH shrinkage check. Woodland species are skipped when their codes are entered in Woodland species to exclude from shrinking diameters: $woodlandText."
        }
        "Height remeasurement check" {
            return "This finding compares the current tree measurement with the prior period. It flags TotalHeight growth over $heightJump and height decreases on live/new live trees."
        }
        "TreeHistory transition check" {
            return "This finding means the current TreeHistory conflicts with the prior measurement, such as old mortality returning to live, live/missed moving directly to old mortality/harvest, or repeated mortality/harvest/thin states. Reference if useful: " + (Get-TreeHistoryClassCrosswalkSummary)
        }
        "Lazarus tree check" {
            return "This finding means a tree previously coded mortality, harvest, thinned, old mortality/harvest, or similar later returned to live, missed, or new ingrowth. TreeHistory 9 include/non-include corrections are allowed and are not flagged. Verify the field record before changing status. Reference if useful: " + (Get-TreeHistoryClassCrosswalkSummary)
        }
        "Ingrowth on install plot" {
            return "This finding means TreeHistory 10 was recorded on an install plot with PlotStatus 2. New trees on install plots should normally be TreeHistory 0; TreeHistory 10 is for new trees on remeasurement plots."
        }
        "TreeHistory 10 in initial project period" {
            return "This finding means TreeHistory 10 was recorded in the first project measurement period. The ingrowth guide says first-period ingrowth records should be corrected to TreeHistory 0."
        }
        "Ingrowth with prior tree measurement" {
            return "This finding means TreeHistory 10 was recorded even though the tree has an earlier nonblank TreeHistory value. Earlier blank TreeHistory rows do not count as prior history for this rule."
        }
        "Possible missed ingrowth" {
            return "This finding means a first-recorded live TreeHistory 0 row was found on a remeasurement plot with PlotStatus 1. Review whether it should be TreeHistory 10. Period 1 and install-plot trees should remain TreeHistory 0."
        }
        "TreeClass/TreeHistory conversion mismatch" {
            return "This finding means the recorded TreeClass and TreeHistory do not match a clear mapping in the supplied crosswalk. Do not guess the correction; verify the original source record, problem codes, and previous TreeHistory before changing either value. Reference: " + (Get-TreeHistoryClassCrosswalkSummary)
        }
        "TreeClass/TreeHistory conversion review" {
            return "This finding means the recorded TreeClass and TreeHistory match a conditional crosswalk case. DADA flags it for hand review because class 0, class 10, and some mortality classes can depend on problem codes, snag/downed status, previous history, or project use. TreeClass 9 is treated as live for current checks. Reference: " + (Get-TreeHistoryClassCrosswalkSummary)
        }
        "TreeStatus/TreeHistory conversion mismatch" {
            return "This finding means the recorded TreeStatus and TreeHistory do not match a clear mapping in the supplied crosswalk. Verify the source status and TreeHistory before editing. Reference: " + (Get-TreeHistoryClassCrosswalkSummary)
        }
        "TreeStatus/TreeHistory conversion review" {
            return "This finding means the recorded TreeStatus and TreeHistory match a conditional crosswalk case. DADA flags it for hand review because status 4 and status 9 can depend on snag/downed status, prior TreeHistory, or project use. Reference: " + (Get-TreeHistoryClassCrosswalkSummary)
        }
        "IDBH shrinkage check" {
            return "This finding flags selected current-period IDBH smaller than the selected previous-period IDBH only for live timber trees. Woodland species are allowed to shrink and should be added to Woodland species to exclude from shrinking diameters, not forced upward, when applicable. Current woodland species exclusion option: $woodlandText."
        }
        "Height shrinkage check" {
            return "This finding flags later TotalHeight smaller than earlier TotalHeight for all-period pairs that are not already reported by the normal previous/current remeasurement check. It may be a data entry issue, wrong tree, broken top, measurement method change, or true field condition. Live-to-dead height shrinkage is not flagged when the later TreeHistory is a mortality status selected in Run options as requiring height, because dead trees can lose height through breakage."
        }
        "Problem severity mismatch" {
            return "This finding means the paired problem and severity fields disagree. Blank, Null, and 0 count as none; values greater than 0 count as entered."
        }
        "New mortality Severity1 check" {
            return "This finding means TreeHistory 2 or 3 was recorded with Problem1 entered, but Severity1 is not code 3. Confirm the new mortality/problem code, then set Severity1 to 3 or correct TreeHistory/Problem1."
        }
        "Missing tree measurement" {
            return "This finding means a tree has at least one measurement but is missing an expected TreeMeasurements row for another project period."
        }
        "Missing plot measurement" {
            return "This finding means a plot has at least one measurement but is missing an expected PlotMeasurements row for another project period."
        }
        "Duplicate selected key" {
            return "This finding means more than one row shares the selected plot/tree/regen key field combination."
        }
        "Workbook count check" {
            return "This finding comes from the cleaning workbook count checks. Custom measurement tables should line up with their parent measurement rows for the relevant periods."
        }
    }

    return "This finding should be interpreted using the coded app rules, the project template metadata, and the uploaded project manual."
}

function Get-TemplateCodeContextForField {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [string]$FieldName,
        [int]$MaxCodes = 80
    )

    if ([string]::IsNullOrWhiteSpace($TableName) -or [string]::IsNullOrWhiteSpace($FieldName)) { return "" }
    if ($FieldName -match "/") { return "" }
    if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $TableName -FieldName $FieldName)) {
        return "Template metadata: $TableName.$FieldName is not marked Active in AppColumns for data-entry review."
    }

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not ($tables -contains "AppTables") -or
        -not ($tables -contains "AppColumns") -or
        -not ($tables -contains "AppColumnCodes")) {
        return "Template metadata: AppTables/AppColumns/AppColumnCodes were not found in this database."
    }

    $metadataSql = @"
SELECT c.[ID], c.[ColumnName], c.[CategoryTypeID]
FROM [AppTables] AS t
INNER JOIN [AppColumns] AS c ON t.[ID] = c.[TableID]
WHERE t.[TableName] = $(Sql-Text $TableName)
  AND c.[ColumnName] = $(Sql-Text $FieldName)
"@
    $metadata = Get-DataTable -Connection $Connection -Sql $metadataSql
    if ($metadata.Rows.Count -eq 0) {
        return "Template metadata: no AppColumns entry was found for $TableName.$FieldName."
    }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine("Template metadata for $($TableName).$($FieldName):")
    foreach ($row in $metadata.Rows) {
        $columnId = [string]$row["ID"]
        $categoryTypeId = [string]$row["CategoryTypeID"]
        [void]$builder.AppendLine("- AppColumns ID $columnId, CategoryTypeID $categoryTypeId.")

        $codes = Get-DataTable -Connection $Connection -Sql "SELECT [Code], [CodeLabel] FROM [AppColumnCodes] WHERE [ColumnID] = $(Sql-Text $columnId)"
        if ($codes.Rows.Count -eq 0) {
            [void]$builder.AppendLine("- No AppColumnCodes rows were found for this field.")
            continue
        }

        $validCodes = New-Object System.Collections.Generic.List[string]
        $codeExamples = New-Object System.Collections.Generic.List[string]
        foreach ($codeRow in $codes.Rows) {
            $code = Normalize-CodeValue $codeRow["Code"]
            Add-UniqueCode -Codes $validCodes -Code $code
            if ($code -match "^\d+$") {
                Add-UniqueCode -Codes $validCodes -Code (([int64]$code).ToString())
            }

            $label = Normalize-CodeValue $codeRow["CodeLabel"]
            if (-not [string]::IsNullOrWhiteSpace($code) -and -not [string]::IsNullOrWhiteSpace($label)) {
                Add-UniqueCode -Codes $codeExamples -Code "$code = $label"
            }
        }

        $displayValues = @()
        if ($codeExamples.Count -gt 0) {
            $displayValues = @($codeExamples.ToArray() | Sort-Object -Unique)
        }
        else {
            $displayValues = @($validCodes.ToArray() | Sort-Object -Unique)
        }

        if ($displayValues.Count -gt 0) {
            $shownValues = @($displayValues | Select-Object -First $MaxCodes)
            $prefix = if ($displayValues.Count -gt $shownValues.Count) { "Valid codes, showing $($shownValues.Count) of $($displayValues.Count): " } else { "Valid codes: " }
            [void]$builder.AppendLine("- $prefix$([string]::Join('; ', [string[]]$shownValues))")
        }
    }

    return Limit-TextLength -Text $builder.ToString() -MaxLength 3500
}

function Get-TemplateCodeFieldSummaryForAi {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [int]$MaxFields = 40
    )

    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not ($tables -contains "AppTables") -or
        -not ($tables -contains "AppColumns") -or
        -not ($tables -contains "AppColumnCodes")) {
        return "Template code metadata summary: AppTables/AppColumns/AppColumnCodes were not found."
    }

    $targetTables = @(Get-CfiWorkbookTargetTables -Connection $Connection)
    $targetSql = if ($targetTables.Count -gt 0) {
        " AND t.[TableName] In (" + (($targetTables | ForEach-Object { Sql-Text $_ }) -join ", ") + ")"
    }
    else {
        ""
    }

    $metadataSql = @"
SELECT t.[TableName], c.[ID], c.[ColumnName]
FROM [AppTables] AS t
INNER JOIN [AppColumns] AS c ON t.[ID] = c.[TableID]
WHERE c.[CategoryTypeID] = 3$targetSql
ORDER BY t.[TableName], c.[ColumnName]
"@
    $metadata = Get-DataTable -Connection $Connection -Sql $metadataSql
    if ($metadata.Rows.Count -eq 0) {
        return "Template code metadata summary: no coded AppColumns fields were found for the cleaning target tables."
    }

    $codesTable = Get-DataTable -Connection $Connection -Sql "SELECT [ColumnID], [Code], [CodeLabel] FROM [AppColumnCodes]"
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine("Template code metadata summary for AI:")

    $fieldIndex = 0
    foreach ($row in $metadata.Rows) {
        if ($fieldIndex -ge $MaxFields) { break }
        $tableName = Get-DataRowText -Row $row -ColumnName "TableName"
        $fieldName = Get-DataRowText -Row $row -ColumnName "ColumnName"
        if (-not (Test-AppColumnFieldActiveForReview -Connection $Connection -TableName $tableName -FieldName $fieldName)) { continue }
        $fieldIndex++
        $columnId = [string]$row["ID"]
        $examples = New-Object System.Collections.Generic.List[string]
        $codeCount = 0
        foreach ($codeRow in $codesTable.Rows) {
            if ([string]$codeRow["ColumnID"] -ne $columnId) { continue }
            $codeCount++
            if ($examples.Count -ge 5) { continue }
            $code = Normalize-CodeValue $codeRow["Code"]
            $label = Normalize-CodeValue $codeRow["CodeLabel"]
            if ([string]::IsNullOrWhiteSpace($code)) { continue }
            if ([string]::IsNullOrWhiteSpace($label)) {
                Add-UniqueCode -Codes $examples -Code $code
            }
            else {
                Add-UniqueCode -Codes $examples -Code "$code = $label"
            }
        }

        $exampleText = if ($examples.Count -gt 0) { "; examples: " + [string]::Join("; ", [string[]]$examples.ToArray()) } else { "" }
        [void]$builder.AppendLine(("- {0}.{1}: {2} valid code(s){3}" -f $tableName, $fieldName, $codeCount, $exampleText))
    }

    if ($fieldIndex -eq 0) {
        [void]$builder.AppendLine("- No active coded AppColumns fields were found for the cleaning target tables.")
    }
    elseif ($metadata.Rows.Count -gt $MaxFields) {
        [void]$builder.AppendLine("- Additional coded fields omitted from chat context to keep the AI request compact.")
    }

    return Limit-TextLength -Text $builder.ToString() -MaxLength 7000
}

function Get-TemplateRuleContextForAuditRow {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.DataRow]$Row
    )

    try {
        $tableName = Get-DataRowText -Row $Row -ColumnName "TableName"
        $fieldName = Get-DataRowText -Row $Row -ColumnName "FieldName"
        $ruleName = Get-DataRowText -Row $Row -ColumnName "RuleName"

        $builder = New-Object System.Text.StringBuilder
        [void]$builder.AppendLine((Get-SpecificCodedRuleContext -TableName $tableName -FieldName $fieldName -RuleName $ruleName))
        [void]$builder.AppendLine("")
        [void]$builder.AppendLine((Get-CodedCleaningRulesSummary))

        $templateContext = Get-TemplateCodeContextForField -Connection $Connection -TableName $tableName -FieldName $fieldName
        if (-not [string]::IsNullOrWhiteSpace($templateContext)) {
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine($templateContext)
        }

        $fieldTip = Get-FieldManualTip -TableName $tableName -FieldName $fieldName
        if (-not [string]::IsNullOrWhiteSpace($fieldTip)) {
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine("Cleaning workbook/field tip for this field:")
            [void]$builder.AppendLine((Limit-TextLength -Text $fieldTip -MaxLength 1200))
        }

        return Limit-TextLength -Text $builder.ToString() -MaxLength 7000
    }
    catch {
        return "Coded/template rule context could not be built for this finding: $($_.Exception.Message)"
    }
}

function Get-CleaningRulesContextForAi {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine((Get-CodedCleaningRulesSummary))
    [void]$builder.AppendLine("")
    try {
        [void]$builder.AppendLine((Get-TemplateCodeFieldSummaryForAi -Connection $Connection))
    }
    catch {
        [void]$builder.AppendLine("Template code metadata summary could not be read: $($_.Exception.Message)")
    }

    return Limit-TextLength -Text $builder.ToString() -MaxLength 9000
}

function Get-TreePreviousMeasurementFromSql {
    param(
        [switch]$IncludeTreeMetadata,
        [switch]$IncludePlotMetadata
    )

    if ($IncludePlotMetadata) {
        return @"
FROM
    ((([TreeMeasurements] AS cur
    INNER JOIN (
        SELECT c.[TreeKey], c.[PeriodNumber], Max(p.[PeriodNumber]) AS [PrevPeriodNumber]
        FROM [TreeMeasurements] AS c
        INNER JOIN [TreeMeasurements] AS p
            ON c.[TreeKey] = p.[TreeKey]
           AND p.[PeriodNumber] < c.[PeriodNumber]
        GROUP BY c.[TreeKey], c.[PeriodNumber]
    ) AS link
        ON cur.[TreeKey] = link.[TreeKey]
       AND cur.[PeriodNumber] = link.[PeriodNumber])
    INNER JOIN [TreeMeasurements] AS prev
        ON prev.[TreeKey] = link.[TreeKey]
       AND prev.[PeriodNumber] = link.[PrevPeriodNumber])
    LEFT JOIN [Trees] AS treeMeta
        ON cur.[TreeID] = treeMeta.[TreeID])
    LEFT JOIN [Plots] AS plotMeta
        ON treeMeta.[PlotID] = plotMeta.[PlotID]
"@
    }

    if ($IncludeTreeMetadata) {
        return @"
FROM
    (([TreeMeasurements] AS cur
    INNER JOIN (
        SELECT c.[TreeKey], c.[PeriodNumber], Max(p.[PeriodNumber]) AS [PrevPeriodNumber]
        FROM [TreeMeasurements] AS c
        INNER JOIN [TreeMeasurements] AS p
            ON c.[TreeKey] = p.[TreeKey]
           AND p.[PeriodNumber] < c.[PeriodNumber]
        GROUP BY c.[TreeKey], c.[PeriodNumber]
    ) AS link
        ON cur.[TreeKey] = link.[TreeKey]
       AND cur.[PeriodNumber] = link.[PeriodNumber])
    INNER JOIN [TreeMeasurements] AS prev
        ON prev.[TreeKey] = link.[TreeKey]
       AND prev.[PeriodNumber] = link.[PrevPeriodNumber])
    LEFT JOIN [Trees] AS treeMeta
        ON cur.[TreeID] = treeMeta.[TreeID]
"@
    }

    return @"
FROM
    ([TreeMeasurements] AS cur
    INNER JOIN (
        SELECT c.[TreeKey], c.[PeriodNumber], Max(p.[PeriodNumber]) AS [PrevPeriodNumber]
        FROM [TreeMeasurements] AS c
        INNER JOIN [TreeMeasurements] AS p
            ON c.[TreeKey] = p.[TreeKey]
           AND p.[PeriodNumber] < c.[PeriodNumber]
        GROUP BY c.[TreeKey], c.[PeriodNumber]
    ) AS link
        ON cur.[TreeKey] = link.[TreeKey]
       AND cur.[PeriodNumber] = link.[PeriodNumber])
    INNER JOIN [TreeMeasurements] AS prev
        ON prev.[TreeKey] = link.[TreeKey]
       AND prev.[PeriodNumber] = link.[PrevPeriodNumber]
"@
}

function Get-TreeMeasurementPairFromSql {
    param(
        [switch]$IncludeTreeMetadata,
        [switch]$IncludePlotMetadata
    )

    if ($IncludePlotMetadata) {
        return @"
FROM (([TreeMeasurements] AS earlier
INNER JOIN [TreeMeasurements] AS later
    ON earlier.[TreeID] = later.[TreeID])
LEFT JOIN [Trees] AS treeMeta
    ON later.[TreeID] = treeMeta.[TreeID])
LEFT JOIN [Plots] AS plotMeta
    ON treeMeta.[PlotID] = plotMeta.[PlotID]
"@
    }

    if ($IncludeTreeMetadata) {
        return @"
FROM ([TreeMeasurements] AS earlier
INNER JOIN [TreeMeasurements] AS later
    ON earlier.[TreeID] = later.[TreeID])
LEFT JOIN [Trees] AS treeMeta
    ON later.[TreeID] = treeMeta.[TreeID]
"@
    }

    return @"
FROM [TreeMeasurements] AS earlier
INNER JOIN [TreeMeasurements] AS later
    ON earlier.[TreeID] = later.[TreeID]
"@
}

function Get-VerificationSqlForAuditRow {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.DataRow]$Row
    )

    $tableName = [string]$Row["TableName"]
    $ruleName = [string]$Row["RuleName"]
    $fieldName = [string]$Row["FieldName"]
    $recordLabel = [string]$Row["RecordLabel"]
    if ([string]::IsNullOrWhiteSpace($tableName)) { return "" }

    $table = Quote-Name $tableName
    $field = if (-not [string]::IsNullOrWhiteSpace($fieldName) -and $fieldName -notmatch "/") { Quote-Name $fieldName } else { $null }
    $idbhMaxSql = (Get-ControlDecimalValue -Control $dbhMax -DefaultValue 500).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $heightMaxSql = (Get-ControlDecimalValue -Control $heightMax -DefaultValue 150).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $stemMaxSql = (Get-StemCountMaximumValue).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $idbhJumpSql = (Get-ControlDecimalValue -Control $dbhGrowthMax -DefaultValue 100).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $heightJumpSql = (Get-ControlDecimalValue -Control $heightGrowthMax -DefaultValue 20).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $treeMetadataAvailability = [pscustomobject]@{
        IncludeTreeMetadata = $false
        IncludePlotMetadata = $false
    }
    try {
        $allTables = @(Get-UserTables -Connection $Connection -IncludeAudit)
        if (Test-TableAvailable -Tables $allTables -TableName "TreeMeasurements") {
            $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
            $treeMetadataAvailability = Get-TreeMetadataAvailability -Connection $Connection -Tables $allTables -MeasurementColumns $treeMeasurementColumns
        }
    }
    catch {
        $treeMetadataAvailability = [pscustomobject]@{
            IncludeTreeMetadata = $false
            IncludePlotMetadata = $false
        }
    }
    $treeAlias = if ($treeMetadataAvailability.IncludeTreeMetadata) { "treeMeta" } else { "" }
    $plotAlias = if ($treeMetadataAvailability.IncludePlotMetadata) { "plotMeta" } else { "" }
    if ($null -eq $treeMeasurementColumns) { $treeMeasurementColumns = @() }
    $treeColumnsForLiveStatus = @()
    if ($treeMetadataAvailability.IncludeTreeMetadata) {
        try {
            $treeColumnsForLiveStatus = @(Get-TableColumns -Connection $Connection -TableName "Trees")
        }
        catch {
            $treeColumnsForLiveStatus = @()
        }
    }
    $dbhShrinkConditionForPrevious = Get-LiveTimberDbhShrinkSqlCondition `
        -CurrentAlias "cur" `
        -PreviousAlias "prev" `
        -TreeAlias $treeAlias `
        -PlotAlias $plotAlias
    $dbhShrinkConditionForPair = Get-LiveTimberDbhShrinkSqlCondition `
        -CurrentAlias "later" `
        -PreviousAlias "earlier" `
        -TreeAlias $treeAlias `
        -PlotAlias $plotAlias

    switch ($ruleName) {
        "Regen IDBH code check" {
            $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements")
            if (-not (Test-ColumnExists -Columns $regenColumns -Name "SpeciesCode") -or
                -not (Test-ColumnExists -Columns $regenColumns -Name "IDBH")) { return "" }
            $condition = Get-RegenIdbhRequiredCondition -SpeciesField "[SpeciesCode]" -IdbhField "[IDBH]"
            return New-ScopedSelectVerificationSql -Connection $Connection -TableName "RegenMeasurements" -RuleName $ruleName -FieldName "IDBH" -Condition $condition
        }
        "Required regen minor plot" {
            return Get-RequiredRegenMinorPlotVerificationSql -Connection $Connection
        }
        "Regen minor plot check" {
            return Get-RegenMinorPlotVerificationSql -Connection $Connection
        }
        "Rule coverage diagnostic" {
            if ($tableName -eq "RegenMeasurements") {
                return "SELECT Count(*) AS [RegenRows] FROM [RegenMeasurements];"
            }
            return "SELECT Count(*) AS [RowsInTable] FROM $table;"
        }
        "Required plot entry" {
            return Get-RequiredPlotEntryVerificationSql -Connection $Connection -TableName $tableName -FieldName $fieldName
        }
        "Required plot remark" {
            return Get-PlotStatusRemarkVerificationSql -Connection $Connection -TableName $tableName
        }
        "PlotStatus progression check" {
            return Get-PlotStatusProgressionVerificationSql -Connection $Connection
        }
        "Plot data on not-measured plot" {
            return Get-NoPlotDataOnInactivePlotStatusVerificationSql -Connection $Connection -TableName $tableName -FieldName $fieldName
        }
        "Regen data on not-measured plot" {
            return Get-NoRegenDataOnInactivePlotStatusVerificationSql -Connection $Connection -TableName $tableName -FieldName $fieldName
        }
        "DMUML code check" {
            if ($null -eq $field) { return "" }
            $condition = Get-DmuMlAuditCondition -QuotedField $field
            return New-ScopedSelectVerificationSql -Connection $Connection -TableName $tableName -RuleName $ruleName -FieldName $fieldName -Condition $condition
        }
        "Invalid CFI code" {
            if ($null -eq $field) { return "" }
            if (Test-DmuMlField -FieldName $fieldName) {
                $condition = Get-DmuMlAuditCondition -QuotedField $field
                return New-ScopedSelectVerificationSql -Connection $Connection -TableName $tableName -RuleName $ruleName -FieldName $fieldName -Condition $condition
            }
            $validList = Get-ValidCodesSqlList -Connection $Connection -TableName $tableName -FieldName $fieldName
            $condition = if ([string]::IsNullOrWhiteSpace($validList)) { "Trim(($field & '')) <> ''" } else { "Trim(($field & '')) <> '' AND Trim(($field & '')) Not In ($validList)" }
            return New-ScopedSelectVerificationSql -Connection $Connection -TableName $tableName -RuleName $ruleName -FieldName $fieldName -Condition $condition
        }
        "Missing DBH" {
            if ($tableName -eq "TreeMeasurements" -and $fieldName -eq "IDBH") {
                $missingRequiredCondition = Get-TreeIdbhMissingRequiredCondition -TreeHistoryField "[TreeHistory]" -IdbhField "[IDBH]"
                return New-ScopedSelectVerificationSql -Connection $Connection -TableName "TreeMeasurements" -RuleName $ruleName -FieldName "IDBH" -Condition $missingRequiredCondition
            }
            return ""
        }
        "DBH should be blank" {
            if ($tableName -eq "TreeMeasurements" -and $fieldName -eq "IDBH") {
                $shouldBeBlankCondition = Get-TreeIdbhShouldBeBlankCondition -TreeHistoryField "[TreeHistory]" -IdbhField "[IDBH]"
                return New-ScopedSelectVerificationSql -Connection $Connection -TableName "TreeMeasurements" -RuleName $ruleName -FieldName "IDBH" -Condition $shouldBeBlankCondition
            }
            return ""
        }
        "Required regen species" {
            return Get-RequiredRegenSpeciesVerificationSql -Connection $Connection
        }
        "Required tree species" {
            return Get-RequiredTreeSpeciesVerificationSql -Connection $Connection
        }
        "Required crown ratio" {
            return Get-RequiredTreeCrownVerificationSql -Connection $Connection -FieldName $fieldName
        }
        "Required crown class" {
            return Get-RequiredTreeCrownVerificationSql -Connection $Connection -FieldName $fieldName
        }
        "Required radial increment" {
            return Get-RequiredTreeRadialIncrementVerificationSql -Connection $Connection -FieldName $fieldName
        }
        "Tree data on not-measured plot" {
            return Get-NoTreeDataOnInactivePlotStatusVerificationSql -Connection $Connection -TableName $tableName -FieldName $fieldName
        }
        "IDBH range check" {
            $targetField = if ($null -ne $field) { $field } else { "[IDBH]" }
            if ($tableName -eq "RegenMeasurements" -and $fieldName -eq "IDBH") {
                $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements")
                if (-not (Test-ColumnExists -Columns $regenColumns -Name "SpeciesCode") -or
                    -not (Test-ColumnExists -Columns $regenColumns -Name "IDBH")) { return "" }
                $condition = Get-RegenIdbhRequiredCondition -SpeciesField "[SpeciesCode]" -IdbhField "[IDBH]"
                return New-ScopedSelectVerificationSql -Connection $Connection -TableName "RegenMeasurements" -RuleName "Regen IDBH code check" -FieldName "IDBH" -Condition $condition
            }
            if ($tableName -eq "TreeMeasurements" -and $fieldName -eq "IDBH") {
                return New-ScopedSelectVerificationSql -Connection $Connection -TableName "TreeMeasurements" -RuleName $ruleName -FieldName "IDBH" -Condition "[IDBH] Is Not Null AND ([IDBH] <= 0 OR [IDBH] > $idbhMaxSql)"
            }
            return New-ScopedSelectVerificationSql -Connection $Connection -TableName $tableName -RuleName $ruleName -FieldName $fieldName -Condition "$targetField Is Null OR $targetField <= 0 OR $targetField > $idbhMaxSql"
        }
        "Height range check" {
            $targetField = if ($null -ne $field) { $field } else { "[TotalHeight]" }
            if ($tableName -eq "TreeMeasurements" -and $fieldName -eq "TotalHeight") {
                return Get-TreeHeightVerificationSql -Connection $Connection
            }
            return New-ScopedSelectVerificationSql -Connection $Connection -TableName $tableName -RuleName $ruleName -FieldName $fieldName -Condition "$targetField Is Not Null AND ($targetField <= 0 OR $targetField > $heightMaxSql)"
        }
        "TotalHeight subsample review" {
            return Get-TreeHeightSubsampleVerificationSql -Connection $Connection
        }
        "TotalHeight subsample order review" {
            return Get-TreeHeightSubsampleOrderVerificationSql -Connection $Connection
        }
        "Stem count range check" {
            $targetField = if ($null -ne $field) { $field } else { "[StemCount]" }
            if ($tableName -eq "RegenMeasurements" -and $fieldName -eq "StemCount") {
                $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements")
                if (-not (Test-ColumnExists -Columns $regenColumns -Name "SpeciesCode") -or
                    -not (Test-ColumnExists -Columns $regenColumns -Name "StemCount")) { return "" }
                $condition = Get-RegenStemCountRequiredCondition -SpeciesField "[SpeciesCode]" -StemCountField "[StemCount]" -MaximumValue (Get-StemCountMaximumValue)
                return New-ScopedSelectVerificationSql -Connection $Connection -TableName "RegenMeasurements" -RuleName $ruleName -FieldName "StemCount" -Condition $condition
            }
            return New-ScopedSelectVerificationSql -Connection $Connection -TableName $tableName -RuleName $ruleName -FieldName $fieldName -Condition "$targetField Is Null OR $targetField <= 0 OR $targetField > $stemMaxSql"
        }
        "Tree stem count check" {
            $woodlandList = Get-WoodlandSpeciesCodesSqlList
            $canJoinSpecies = $false
            try {
                $tablesForTreeStem = @(Get-UserTables -Connection $Connection -IncludeAudit)
                $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
                if ((Test-TableAvailable -Tables $tablesForTreeStem -TableName "Trees") -and
                    (Test-ColumnExists -Columns $treeMeasurementColumns -Name "TreeID")) {
                    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
                    $canJoinSpecies =
                        (Test-ColumnExists -Columns $treeColumns -Name "TreeID") -and
                        (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode")
                }
            }
            catch {
                $canJoinSpecies = $false
            }

            if ($canJoinSpecies) {
                $conditions = New-Object System.Collections.Generic.List[string]
                [void]$conditions.Add("(tm.[StemCount] Is Not Null AND (tm.[StemCount] <= 0 OR tm.[StemCount] > $stemMaxSql))")
                if (-not [string]::IsNullOrWhiteSpace($woodlandList)) {
                    [void]$conditions.Add("(Trim((treeMeta.[SpeciesCode] & '')) In ($woodlandList) AND tm.[StemCount] Is Null)")
                }
                $conditionText = [string]::Join(" OR ", [string[]]$conditions.ToArray())
                $conditionText = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $conditionText -Alias "tm"
                $tmSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "tm"
                return @"
SELECT $tmSelect, treeMeta.[SpeciesCode] AS [TreeSpeciesCode]
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS treeMeta
    ON tm.[TreeID] = treeMeta.[TreeID]
WHERE $conditionText;
"@
            }

            return New-ScopedSelectVerificationSql -Connection $Connection -TableName "TreeMeasurements" -RuleName $ruleName -FieldName "StemCount" -Condition "[StemCount] Is Not Null AND ([StemCount] <= 0 OR [StemCount] > $stemMaxSql)"
        }
        "IDBH jump check" {
            $fromSql = Get-TreePreviousMeasurementFromSql `
                -IncludeTreeMetadata:($treeMetadataAvailability.IncludeTreeMetadata) `
                -IncludePlotMetadata:($treeMetadataAvailability.IncludePlotMetadata)
            $extraColumns = ""
            if ($treeMetadataAvailability.IncludeTreeMetadata) { $extraColumns += ", treeMeta.[SpeciesCode] AS [TreeSpeciesCode]" }
            $whereClause = "cur.[IDBH] Is Not Null AND prev.[IDBH] Is Not Null AND (cur.[IDBH] - prev.[IDBH]) > $idbhJumpSql"
            $currentOnlyCondition = Get-CurrentMeasurementPeriodAliasCondition -Alias "cur"
            if (-not [string]::IsNullOrWhiteSpace($currentOnlyCondition)) {
                $whereClause = "(($whereClause) AND ($currentOnlyCondition))"
            }
            $selectedRemeasurementPairCondition = Get-SelectedPreviousCurrentPairCondition -EarlierAlias "prev" -LaterAlias "cur"
            if (-not [string]::IsNullOrWhiteSpace($selectedRemeasurementPairCondition)) {
                $whereClause = "(($whereClause) AND ($selectedRemeasurementPairCondition))"
            }
            $curSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "cur"
            return @"
SELECT $curSelect, prev.[PeriodNumber] AS [PreviousPeriodNumber], prev.[IDBH] AS [PreviousIDBH], (cur.[IDBH] - prev.[IDBH]) AS [IDBHJump]$extraColumns
$fromSql
WHERE $whereClause;
"@
        }
        "Height remeasurement check" {
            $fromSql = Get-TreePreviousMeasurementFromSql `
                -IncludeTreeMetadata:($treeMetadataAvailability.IncludeTreeMetadata) `
                -IncludePlotMetadata:($treeMetadataAvailability.IncludePlotMetadata)
            $heightCurrentLiveCondition = Get-TreeHeightStatusSqlCondition `
                -MeasurementColumns $treeMeasurementColumns `
                -TreeColumns $treeColumnsForLiveStatus `
                -MeasurementAlias "cur" `
                -TreeAlias $treeAlias `
                -TreeHistoryCodes @("0", "5", "10")
            $whereClause = "cur.[TotalHeight] Is Not Null AND prev.[TotalHeight] Is Not Null AND ((cur.[TotalHeight] - prev.[TotalHeight]) > $heightJumpSql OR (($heightCurrentLiveCondition) AND cur.[TotalHeight] < prev.[TotalHeight]))"
            $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "cur"
            $curSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "cur"
            return @"
SELECT $curSelect, prev.[PeriodNumber] AS [PreviousPeriodNumber], prev.[TotalHeight] AS [PreviousTotalHeight], (cur.[TotalHeight] - prev.[TotalHeight]) AS [HeightChange]
$fromSql
WHERE $whereClause;
"@
        }
        "TreeHistory transition check" {
            $fromSql = Get-TreePreviousMeasurementFromSql
            $curSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "cur"
            return @"
SELECT $curSelect, prev.[PeriodNumber] AS [PreviousPeriodNumber], prev.[TreeHistory] AS [PreviousTreeHistory]
$fromSql
WHERE cur.[TreeHistory] Is Not Null
  AND prev.[TreeHistory] Is Not Null
  AND (
        (prev.[TreeHistory] In (1, 2, 3, 4, 7, 8) AND cur.[TreeHistory] = 0)
        OR (prev.[TreeHistory] In (0, 10) AND cur.[TreeHistory] In (7, 8))
        OR (prev.[TreeHistory] In (1, 2, 3, 4) AND cur.[TreeHistory] In (1, 2, 3, 4))
        OR (prev.[TreeHistory] In (7, 8) AND cur.[TreeHistory] In (1, 2, 3, 4))
      );
"@
        }
        "Lazarus tree check" {
            $laterSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "later"
            return @"
SELECT $laterSelect, earlier.[PeriodNumber] AS [EarlierPeriod], earlier.[TreeHistory] AS [EarlierTreeHistory]
FROM [TreeMeasurements] AS earlier
INNER JOIN [TreeMeasurements] AS later
    ON earlier.[TreeID] = later.[TreeID]
WHERE earlier.[TreeHistory] In (1, 2, 3, 4, 7, 8)
  AND later.[PeriodNumber] > earlier.[PeriodNumber]
  AND later.[TreeHistory] In (0, 5, 10);
"@
        }
        "Ingrowth on install plot" {
            return Get-IngrowthVerificationSql -Connection $Connection -RuleName $ruleName
        }
        "TreeHistory 10 in initial project period" {
            return Get-IngrowthVerificationSql -Connection $Connection -RuleName $ruleName
        }
        "Ingrowth with prior tree measurement" {
            return Get-IngrowthVerificationSql -Connection $Connection -RuleName $ruleName
        }
        "Possible missed ingrowth" {
            return Get-IngrowthVerificationSql -Connection $Connection -RuleName $ruleName
        }
        "TreeClass/TreeHistory conversion mismatch" {
            return Get-TreeHistoryConversionVerificationSql -Connection $Connection -SourceKind "TreeClass" -Mode "Mismatch"
        }
        "TreeClass/TreeHistory conversion review" {
            return Get-TreeHistoryConversionVerificationSql -Connection $Connection -SourceKind "TreeClass" -Mode "Review"
        }
        "TreeStatus/TreeHistory conversion mismatch" {
            return Get-TreeHistoryConversionVerificationSql -Connection $Connection -SourceKind "TreeStatus" -Mode "Mismatch"
        }
        "TreeStatus/TreeHistory conversion review" {
            return Get-TreeHistoryConversionVerificationSql -Connection $Connection -SourceKind "TreeStatus" -Mode "Review"
        }
        "IDBH shrinkage check" {
            $fromSql = Get-TreeMeasurementPairFromSql `
                -IncludeTreeMetadata:($treeMetadataAvailability.IncludeTreeMetadata) `
                -IncludePlotMetadata:($treeMetadataAvailability.IncludePlotMetadata)
            $extraColumns = ""
            if ($treeMetadataAvailability.IncludeTreeMetadata) { $extraColumns += ", treeMeta.[SpeciesCode] AS [TreeSpeciesCode]" }
            $whereClause = "earlier.[IDBH] Is Not Null AND later.[IDBH] Is Not Null AND later.[PeriodNumber] > earlier.[PeriodNumber] AND $dbhShrinkConditionForPair"
            $selectedPairCondition = Get-SelectedPreviousCurrentPairCondition -EarlierAlias "earlier" -LaterAlias "later"
            if (-not [string]::IsNullOrWhiteSpace($selectedPairCondition)) {
                $whereClause = "(($whereClause) AND ($selectedPairCondition))"
            }
            $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "later"
            $laterSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "later"
            return @"
SELECT $laterSelect, earlier.[PeriodNumber] AS [EarlierPeriod], earlier.[IDBH] AS [EarlierIDBH], (earlier.[IDBH] - later.[IDBH]) AS [Shrinkage]$extraColumns
$fromSql
WHERE $whereClause
ORDER BY (earlier.[IDBH] - later.[IDBH]) DESC;
"@
        }
        "Height shrinkage check" {
            $whereClause = "earlier.[TotalHeight] Is Not Null AND later.[TotalHeight] Is Not Null AND later.[PeriodNumber] > earlier.[PeriodNumber] AND later.[TotalHeight] < earlier.[TotalHeight]"
            $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
            if (Test-ColumnExists -Columns $treeMeasurementColumns -Name "TreeHistory") {
                $deadHeightShrinkAllowedCondition = Get-DeadTreeHeightShrinkAllowedCondition -EarlierAlias "earlier" -LaterAlias "later"
                if (-not [string]::IsNullOrWhiteSpace($deadHeightShrinkAllowedCondition)) {
                    $whereClause = "(($whereClause) AND NOT ($deadHeightShrinkAllowedCondition))"
                }
                $heightShrinkLaterLiveCondition = Get-TreeHeightStatusSqlCondition `
                    -MeasurementColumns $treeMeasurementColumns `
                    -TreeColumns $treeColumnsForLiveStatus `
                    -MeasurementAlias "later" `
                    -TreeAlias $treeAlias `
                    -TreeHistoryCodes @("0", "5", "10")
                $duplicateImmediateHeightShrinkCondition = "NOT (($heightShrinkLaterLiveCondition) AND earlier.[PeriodNumber] = (SELECT Max(priorHeight.[PeriodNumber]) FROM [TreeMeasurements] AS priorHeight WHERE priorHeight.[TreeID] = later.[TreeID] AND priorHeight.[PeriodNumber] < later.[PeriodNumber]))"
                $whereClause = "(($whereClause) AND $duplicateImmediateHeightShrinkCondition)"
            }
            $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause $whereClause -Alias "later"
            $laterSelect = Get-AliasSelectList -Connection $Connection -TableName "TreeMeasurements" -Alias "later"
            $fromSql = Get-TreeMeasurementPairFromSql `
                -IncludeTreeMetadata:($treeMetadataAvailability.IncludeTreeMetadata) `
                -IncludePlotMetadata:($treeMetadataAvailability.IncludePlotMetadata)
            $extraColumns = ""
            if ($treeMetadataAvailability.IncludeTreeMetadata) { $extraColumns += ", treeMeta.[SpeciesCode] AS [TreeSpeciesCode]" }
            return @"
SELECT $laterSelect, earlier.[PeriodNumber] AS [EarlierPeriod], earlier.[TotalHeight] AS [EarlierTotalHeight], (earlier.[TotalHeight] - later.[TotalHeight]) AS [Shrinkage]$extraColumns
$fromSql
WHERE $whereClause
ORDER BY (earlier.[TotalHeight] - later.[TotalHeight]) DESC;
"@
        }
        "Problem severity mismatch" {
            $parts = @($fieldName -split "/" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($parts.Count -eq 2) {
                $condition = Get-ProblemSeverityMismatchCondition -ProblemField $parts[0] -SeverityField $parts[1]
                return New-ScopedSelectVerificationSql -Connection $Connection -TableName "TreeMeasurements" -RuleName $ruleName -FieldName $fieldName -Condition $condition
            }
            $condition1 = Get-ProblemSeverityMismatchCondition -ProblemField "Problem1" -SeverityField "Severity1"
            $condition2 = Get-ProblemSeverityMismatchCondition -ProblemField "Problem2" -SeverityField "Severity2"
            return New-ScopedSelectVerificationSql -Connection $Connection -TableName "TreeMeasurements" -RuleName $ruleName -FieldName $fieldName -Condition "$condition1 OR $condition2"
        }
        "New mortality Severity1 check" {
            $condition = Get-NewMortalityProblem1SeverityCondition
            return New-ScopedSelectVerificationSql -Connection $Connection -TableName "TreeMeasurements" -RuleName $ruleName -FieldName "Severity1" -Condition $condition
        }
        "Missing tree measurement" {
            $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause "NOT EXISTS (
        SELECT 1
        FROM [TreeMeasurements] AS tm
        WHERE tm.[TreeID] = t.[TreeID]
          AND tm.[PeriodNumber] = p.[PeriodNumber]
    )
  AND EXISTS (
        SELECT 1
        FROM [TreeMeasurements] AS tm2
        WHERE tm2.[TreeID] = t.[TreeID]
    )" -Alias "p"
            return @"
SELECT t.[PlotNumber], t.[TreeNumber], p.[PeriodNumber], t.[TreeID]
FROM [Trees] AS t, [ProjectMeasurementPeriods] AS p
WHERE $whereClause
ORDER BY t.[PlotNumber], t.[TreeNumber], p.[PeriodNumber];
"@
        }
        "Missing plot measurement" {
            $whereClause = Add-SelectedPeriodAliasConditionToWhereClause -WhereClause "NOT EXISTS (
        SELECT 1
        FROM [PlotMeasurements] AS meas
        WHERE meas.[PlotID] = pl.[PlotID]
          AND meas.[PeriodNumber] = p.[PeriodNumber]
    )
  AND EXISTS (
        SELECT 1
        FROM [PlotMeasurements] AS meas2
        WHERE meas2.[PlotID] = pl.[PlotID]
    )" -Alias "p"
            return @"
SELECT pl.[PlotNumber], p.[PeriodNumber], pl.[PlotID]
FROM [Plots] AS pl, [ProjectMeasurementPeriods] AS p
WHERE $whereClause
ORDER BY pl.[PlotNumber], p.[PeriodNumber];
"@
        }
        "Duplicate selected key" {
            $fields = @($fieldName -split "/" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($fields.Count -eq 0) { return "" }
            $quotedFields = @($fields | ForEach-Object { Quote-Name $_ })
            $notNullCondition = ($quotedFields | ForEach-Object { "$_ Is Not Null" }) -join " AND "
            $groupBy = $quotedFields -join ", "
            $joinCondition = ($fields | ForEach-Object { "t.$(Quote-Name $_) = d.$(Quote-Name $_)" }) -join " AND "
            $tSelect = Get-AliasSelectList -Connection $Connection -TableName $tableName -Alias "t"
            return "SELECT $tSelect FROM $table AS t INNER JOIN (SELECT $groupBy FROM $table WHERE $notNullCondition GROUP BY $groupBy HAVING Count(*) > 1) AS d ON $joinCondition;"
        }
        "Workbook count check" {
            $periodList = Get-SelectedCleaningPeriodsSqlList
            $directWhere = if ([string]::IsNullOrWhiteSpace($periodList)) { "" } else { " WHERE [PeriodNumber] In ($periodList)" }
            $joinWhere = if ([string]::IsNullOrWhiteSpace($periodList)) { "" } else { " WHERE m.[PeriodNumber] In ($periodList)" }
            switch ($tableName) {
                "PlotMeasurements" { return "SELECT [PeriodNumber], Count(*) AS [MeasurementCount] FROM [PlotMeasurements]$directWhere GROUP BY [PeriodNumber];" }
                "PlotCustomMeasurements" { return "SELECT m.[PeriodNumber], Count(*) AS [CustomMeasurementCount] FROM [PlotCustomMeasurements] AS c INNER JOIN [PlotMeasurements] AS m ON c.[PlotMeasKey] = m.[PlotMeasKey]$joinWhere GROUP BY m.[PeriodNumber];" }
                "TreeMeasurements" { return "SELECT [PeriodNumber], Count(*) AS [MeasurementCount] FROM [TreeMeasurements]$directWhere GROUP BY [PeriodNumber];" }
                "TreeCustomMeasurements" { return "SELECT m.[PeriodNumber], Count(*) AS [CustomMeasurementCount] FROM [TreeCustomMeasurements] AS c INNER JOIN [TreeMeasurements] AS m ON c.[TreeMeasKey] = m.[TreeMeasKey]$joinWhere GROUP BY m.[PeriodNumber];" }
                "RegenCustomMeasurements" {
                    $periodNumber = 0
                    if ($recordLabel -match "Period\s+(\d+)") {
                        [void][int]::TryParse($matches[1], [ref]$periodNumber)
                    }
                    return Get-RegenKeyCountVerificationSql -PeriodNumber $periodNumber
                }
            }
        }
    }

    if ($null -ne $field) {
        return New-ScopedSelectVerificationSql -Connection $Connection -TableName $tableName -RuleName $ruleName -FieldName $fieldName -Condition "$field Is Not Null"
    }

    return "SELECT * FROM $table;"
}

function ConvertTo-VerificationSqlSmokeTestSql {
    param([string]$Sql)

    $trimmed = ([string]$Sql).Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return "" }
    $trimmed = $trimmed.TrimEnd(";").Trim()
    if ($trimmed -match "^\s*SELECT\s+TOP\s+\d+\s") { return "$trimmed;" }
    if ($trimmed -match "\bUNION\b") { return "$trimmed;" }
    if ($trimmed -match "^\s*SELECT\s+DISTINCT\s+") {
        return ([regex]::Replace($trimmed, "^\s*SELECT\s+DISTINCT\s+", "SELECT DISTINCT TOP 1 ", 1)) + ";"
    }
    if ($trimmed -match "^\s*SELECT\s+") {
        return ([regex]::Replace($trimmed, "^\s*SELECT\s+", "SELECT TOP 1 ", 1)) + ";"
    }

    return "$trimmed;"
}

function Assert-VerificationSqlRuns {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$Sql,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Sql)) { return }
    $testSql = ConvertTo-VerificationSqlSmokeTestSql -Sql $Sql
    if ([string]::IsNullOrWhiteSpace($testSql)) { return }

    try {
        [void](Get-DataTable -Connection $Connection -Sql $testSql)
    }
    catch {
        throw "Generated verification SQL failed for $Label. $($_.Exception.Message)"
    }
}

function Remove-VerificationSqlTemporaryColumns {
    param([string]$Sql)

    if ([string]::IsNullOrWhiteSpace($Sql)) { return "" }

    $cleanSql = [string]$Sql
    $replacements = @(
        @{ Pattern = "(?i),\s*[A-Za-z_][A-Za-z0-9_]*\.\[NeedsReview\](?:\s+AS\s+\[[^\]]+\])?"; Replacement = "" },
        @{ Pattern = "(?i),\s*\[NeedsReview\](?:\s+AS\s+\[[^\]]+\])?"; Replacement = "" },
        @{ Pattern = "(?i)(SELECT\s+)[A-Za-z_][A-Za-z0-9_]*\.\[NeedsReview\](?:\s+AS\s+\[[^\]]+\])?\s*,\s*"; Replacement = '$1' },
        @{ Pattern = "(?i)(SELECT\s+)\[NeedsReview\](?:\s+AS\s+\[[^\]]+\])?\s*,\s*"; Replacement = '$1' }
    )

    foreach ($replacement in $replacements) {
        $cleanSql = [regex]::Replace($cleanSql, [string]$replacement.Pattern, [string]$replacement.Replacement)
    }

    return $cleanSql
}

function Assert-VerificationSqlHasNoTemporaryColumns {
    param(
        [string]$Sql,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Sql)) { return }
    if ($Sql -match "(?i)\[NeedsReview\]") {
        throw "Generated verification SQL still contains the temporary NeedsReview column for $Label."
    }
}

function Get-SuggestedEditForAuditRow {
    param([System.Data.DataRow]$Row)

    $ruleName = [string]$Row["RuleName"]
    $fieldName = [string]$Row["FieldName"]

    switch ($ruleName) {
        "DMUML code check" { return "Check the source record. Update DMUML to a valid three-position code using only digits 0, 1, 2, or 3 with a total of 6 or less, such as 000, 001, 020, 200, or 300." }
        "Invalid CFI code" { return "Check the original field record, then replace $fieldName with a valid CFI code from the template metadata. Do not guess when the source is unclear." }
        "Regen IDBH code check" { return "Check the regen tally. If SpeciesCode is correct, update IDBH to one of the valid regen DBH class codes: 0, 20, or 40; otherwise correct SpeciesCode." }
        "Required regen minor plot" { return "Check the regen tally and enter the correct MinorPlot for this regen record, or correct SpeciesCode if the regen row should not be active." }
        "Regen minor plot check" { return "Check the regen tally, species, and IDBH class. Update MinorPlot to the project-required minor plot for that timber/woodland seedling/sapling group, or correct SpeciesCode/IDBH if the group is wrong." }
        "Rule coverage diagnostic" { return "Review the diagnostic text. Update the missing AppColumns active flag, join key, woodland species list, project-specific allowed list, or database field setup, then rerun DADA to confirm the intended checks are covered." }
        "Required plot entry" { return "Check the plot record. Enter the missing $fieldName value if the plot status is 1 or 2, or correct PlotStatus if this plot should not require that field." }
        "Required plot remark" { return "Enter the missing plot remark explaining why PlotStatus is 7, or correct PlotStatus if code 7 was entered by mistake." }
        "PlotStatus progression check" { return "Review the plot's full status history. Correct the first installed period to PlotStatus 2, clear statuses recorded before install, or correct later blank/invalid statuses after confirming the field record." }
        "Plot data on not-measured plot" { return "Clear this plot value for the not-measured/dropped plot period, or correct PlotStatus if the plot was actually measured. Leave plot remarks, UTM coordinates, management unit, and FLCCommercial as needed." }
        "Regen data on not-measured plot" { return "Clear this regen value for the not-measured/dropped plot period, or correct PlotStatus if the plot was actually measured. Leave regen remarks as needed for documentation." }
        "Required regen species" { return "Check the regen tally and enter the correct SpeciesCode, or clear the regen row values if no regen record should have been entered." }
        "Required tree species" { return "Check the tree record and enter the correct SpeciesCode, or confirm the tree measurement row should not exist for the selected period." }
        "Required crown ratio" { return "Check the tree record. Enter the valid crown ratio for TreeHistory 0, 5, or 10, or correct TreeHistory if crown ratio is not required." }
        "Required crown class" { return "Check the tree record. Enter the valid crown class for TreeHistory 0, 5, or 10, or correct TreeHistory if crown class is not required." }
        "Required radial increment" { return "Check the tree record. Enter the radial increment for the timber missed or negative-diameter-growth tree, or correct TreeHistory/species/problem code if radial increment is not required." }
        "Tree data on not-measured plot" { return "Clear this tree value for the not-measured/dropped plot period, or correct PlotStatus if the plot was actually measured." }
        "Missing DBH" { return "Check the source record. Enter the measured DBH, or correct TreeHistory/run options if DBH is not required for this status." }
        "DBH should be blank" { return "Clear IDBH for TreeHistory 1, 4, or 8, or correct TreeHistory if DBH was legitimately measured for this tree." }
        "IDBH range check" { return "Check the source record for a transposed digit or wrong unit. Update IDBH to the measured value." }
        "Height range check" { return "Check the source record and units. Update the height to the measured value, or correct TreeHistory/minor plot if this status and minor plot allow blank height." }
        "Stem count range check" { return "Check the regen tally. Update StemCount to the actual counted stems, or correct/enter SpeciesCode after confirming the regen record." }
        "Tree stem count check" { return "Check the tree tally and species. If the species is woodland/non-timber, update StemCount to the measured stem count. If it is timber, blank StemCount is allowed, but entered values still need to be valid." }
        "IDBH jump check" { return "Compare the prior and current measurements. Correct IDBH if the large increase is a transposed value, wrong tree, or other data-entry issue." }
        "Height remeasurement check" { return "Compare the prior and current measurements. Correct TotalHeight if it was transposed, estimated in the wrong unit, or entered under the wrong tree; otherwise document the true field reason." }
        "TreeHistory transition check" { return "Review the prior and current condition. Update TreeHistory to the field-manual transition that matches the tree's real status." }
        "Lazarus tree check" { return "Review the earlier and later TreeHistory values. Correct the later status if the tree should not return to live/missed/ingrowth after mortality, harvest, thinning, or old mortality/harvest. TreeHistory 9 include/non-include corrections are allowed." }
        "Ingrowth on install plot" { return "Review plot status and source records. If this is a new install plot tree, update TreeHistory from 10 to 0." }
        "TreeHistory 10 in initial project period" { return "Review the source record. If this is the first project measurement period, update TreeHistory from 10 to 0." }
        "Ingrowth with prior tree measurement" { return "Review the earlier nonblank TreeHistory value. If the tree already had a real prior history code, correct TreeHistory 10 to the status that matches the field record." }
        "Possible missed ingrowth" { return "Review whether this first-recorded live tree on a remeasurement plot should be TreeHistory 10 instead of 0." }
        "TreeClass/TreeHistory conversion mismatch" { return "Review the original source record and crosswalk. Correct TreeHistory or TreeClass only after confirming which value was converted incorrectly." }
        "TreeClass/TreeHistory conversion review" { return "Review the original source record, problem codes, previous TreeHistory, and project guidance before deciding whether TreeHistory or TreeClass needs correction." }
        "TreeStatus/TreeHistory conversion mismatch" { return "Review the original source record and crosswalk. Correct TreeHistory or TreeStatus only after confirming which value was converted incorrectly." }
        "TreeStatus/TreeHistory conversion review" { return "Review the original source record, snag/downed status, previous TreeHistory, and project guidance before deciding whether TreeHistory or TreeStatus needs correction." }
        "IDBH shrinkage check" { return "Compare the selected previous/current IDBH values for the live timber tree. Correct IDBH if the shrinkage is a data-entry issue; if it is a woodland species, add its code to Woodland species to exclude from shrinking diameters and rerun." }
        "Height shrinkage check" { return "Compare the earlier/later TotalHeight values for the tree. Correct the height if it is a data-entry issue; live-to-dead shrinkage can be valid when mortality-height collection is enabled because dead trees can lose height through breakage." }
        "Problem severity mismatch" { return "Fill both the problem and severity fields from the source record, or clear both if the issue was entered by mistake." }
        "New mortality Severity1 check" { return "If TreeHistory 2 or 3 and Problem1 are correct, update Severity1 to 3; otherwise correct TreeHistory or Problem1." }
        "Missing tree measurement" { return "Create the missing TreeMeasurements row for the listed tree/period, or confirm the project design intentionally omits that measurement." }
        "Missing plot measurement" { return "Create the missing PlotMeasurements row for the listed plot/period, or confirm the project design intentionally omits that measurement." }
        "Duplicate selected key" { return "Review the duplicate records against the source data. Correct the plot/tree/regen key on the wrong record, or remove only confirmed accidental duplicates." }
        "Workbook count check" { return "Add missing measurement/custom rows or correct period/key links until the count matches the workbook rule for that measurement period." }
        "Project user guide check" { return "Update the project setup item named in the message, then rerun." }
    }

    return "Review the source record, update the field to the verified correct value, and rerun to confirm the finding clears."
}

function Normalize-SuggestedValueText {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) { return "" }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    $text = [regex]::Replace($text, "\s+", " ")
    $text = $text.Trim("`"", "'")
    if ($text -match "^(n/a|na|none|unknown|not enough information|not enough info|cannot determine|do not know)$") { return "" }
    if ($text -match "^(null|nil)$") { return "NULL" }
    if ($text.Length -gt 80) { return "" }
    if ($text -match "[;\r\n]") { return "" }
    return $text
}

function Normalize-AiSuggestedEditValue {
    param([object]$Value)

    $text = Normalize-SuggestedValueText $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    if ($text -match "[,:]") { return "" }
    if ($text -match "[!?]") { return "" }
    if (($text -match "\.\s*$") -and -not ($text -match "^\d+(\.\d+)?$")) { return "" }

    $words = @($text -split "\s+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($words.Count -gt 4) { return "" }
    if ($text -match "(?i)\b(check|update|replace|review|correct|verify|enter|set|change|clear|compare|document|source|record|field|should|likely|maybe|manual|value)\b") { return "" }

    return $text
}

function Get-ObservedAuditValue {
    param(
        [string]$ObservedValue,
        [string]$Prefix,
        [string]$FieldName
    )

    if ([string]::IsNullOrWhiteSpace($ObservedValue) -or [string]::IsNullOrWhiteSpace($FieldName)) { return "" }

    $pattern = [regex]::Escape($Prefix) + "\s+P[^=;]*\s+" + [regex]::Escape($FieldName) + "=([^;]+)"
    if ($ObservedValue -match $pattern) {
        return Normalize-SuggestedValueText $matches[1]
    }

    return ""
}

function Get-ObservedKeyValue {
    param(
        [string]$ObservedValue,
        [string]$KeyName
    )

    if ([string]::IsNullOrWhiteSpace($ObservedValue) -or [string]::IsNullOrWhiteSpace($KeyName)) { return "" }

    $pattern = "(^|;)\s*" + [regex]::Escape($KeyName) + "\s*=\s*([^;]*)"
    if ($ObservedValue -match $pattern) {
        return Normalize-SuggestedValueText $matches[2]
    }

    return ""
}

function Test-DecimalLessThan {
    param(
        [string]$Left,
        [string]$Right
    )

    $leftValue = 0.0
    $rightValue = 0.0
    if (-not [double]::TryParse($Left, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$leftValue)) { return $false }
    if (-not [double]::TryParse($Right, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$rightValue)) { return $false }
    return ($leftValue -lt $rightValue)
}

function Get-SuggestedValueForAuditRow {
    param([System.Data.DataRow]$Row)

    $ruleName = Get-DataRowText -Row $Row -ColumnName "RuleName"
    $observedValue = Get-DataRowText -Row $Row -ColumnName "ObservedValue"

    switch ($ruleName) {
        "IDBH jump check" {
            $previous = Get-ObservedAuditValue -ObservedValue $observedValue -Prefix "Previous" -FieldName "IDBH"
            $current = Get-ObservedAuditValue -ObservedValue $observedValue -Prefix "Current" -FieldName "IDBH"
            if ((Test-DecimalLessThan -Left $current -Right $previous)) { return $previous }
            return ""
        }
        "IDBH shrinkage check" {
            return (Get-ObservedAuditValue -ObservedValue $observedValue -Prefix "Earlier" -FieldName "IDBH")
        }
        "Height remeasurement check" {
            $previous = Get-ObservedAuditValue -ObservedValue $observedValue -Prefix "Previous" -FieldName "TotalHeight"
            $current = Get-ObservedAuditValue -ObservedValue $observedValue -Prefix "Current" -FieldName "TotalHeight"
            if ((Test-DecimalLessThan -Left $current -Right $previous)) { return $previous }
            return ""
        }
        "Height shrinkage check" {
            return (Get-ObservedAuditValue -ObservedValue $observedValue -Prefix "Earlier" -FieldName "TotalHeight")
        }
        "Regen minor plot check" {
            if ($observedValue -match "Expected=([^;]+)$") {
                $values = @($matches[1] -split "," | ForEach-Object { Normalize-SuggestedValueText $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
                if ($values.Count -eq 1) { return [string]$values[0] }
            }
            return ""
        }
    }

    return ""
}

function Test-TreeHistoryRecordedTimelineRow {
    param([System.Data.DataRow]$Row)

    $tableName = Get-DataRowText -Row $Row -ColumnName "TableName"
    $fieldName = Get-DataRowText -Row $Row -ColumnName "FieldName"
    if (-not $tableName.Equals("TreeMeasurements", [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    if (-not $fieldName.Equals("TreeHistory", [System.StringComparison]::OrdinalIgnoreCase)) { return $false }

    return $true
}

function Get-TreeHistoryTimelineColumnCache {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [hashtable]$Cache
    )

    if ($Cache.ContainsKey("__TreeHistoryTimelineColumns")) {
        return [object[]]$Cache["__TreeHistoryTimelineColumns"]
    }

    try {
        $columns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
        $Cache["__TreeHistoryTimelineColumns"] = $columns
        return $columns
    }
    catch {
        $Cache["__TreeHistoryTimelineColumns"] = @()
        return @()
    }
}

function Get-TreeHistoryTimelineRecordedValue {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.DataRow]$Row,
        [hashtable]$Identity,
        [hashtable]$TimelineCache
    )

    if (-not (Test-TreeHistoryRecordedTimelineRow -Row $Row)) { return "" }

    $columns = @(Get-TreeHistoryTimelineColumnCache -Connection $Connection -Cache $TimelineCache)
    foreach ($requiredColumn in @("PeriodNumber", "TreeHistory")) {
        if (-not (Test-ColumnExists -Columns $columns -Name $requiredColumn)) { return "" }
    }

    $sourceRowId = Get-DataRowText -Row $Row -ColumnName "SourceRowId"
    $treeId = ""
    $treeKey = [string]$Identity["TreeKey"]

    if (-not [string]::IsNullOrWhiteSpace($sourceRowId) -and
        (Test-ColumnExists -Columns $columns -Name "MeasurementID")) {
        try {
            $anchorSelects = New-Object System.Collections.Generic.List[string]
            if (Test-ColumnExists -Columns $columns -Name "TreeID") {
                [void]$anchorSelects.Add("([TreeID] & '') AS [TreeID]")
            }
            if (Test-ColumnExists -Columns $columns -Name "TreeKey") {
                [void]$anchorSelects.Add("([TreeKey] & '') AS [TreeKey]")
            }

            if ($anchorSelects.Count -gt 0) {
                $anchor = Get-FirstDataRow -Connection $Connection -Sql ("SELECT " + ([string]::Join(", ", [string[]]$anchorSelects.ToArray())) + " FROM [TreeMeasurements] WHERE ([MeasurementID] & '') = " + (Sql-Text $sourceRowId))
                if ($null -ne $anchor) {
                    $treeId = Get-DataRowText -Row $anchor -ColumnName "TreeID"
                    if ([string]::IsNullOrWhiteSpace($treeKey)) {
                        $treeKey = Get-DataRowText -Row $anchor -ColumnName "TreeKey"
                    }
                }
            }
        }
        catch {
        }
    }

    $whereClause = ""
    $cacheKey = ""
    if (-not [string]::IsNullOrWhiteSpace($treeId) -and
        (Test-ColumnExists -Columns $columns -Name "TreeID")) {
        $whereClause = "([TreeID] & '') = $(Sql-Text $treeId)"
        $cacheKey = "TreeID|$treeId"
    }
    elseif (-not [string]::IsNullOrWhiteSpace($treeKey) -and
        (Test-ColumnExists -Columns $columns -Name "TreeKey")) {
        $whereClause = "([TreeKey] & '') = $(Sql-Text $treeKey)"
        $cacheKey = "TreeKey|$treeKey"
    }
    else {
        return ""
    }

    if ($TimelineCache.ContainsKey($cacheKey)) {
        $cachedTimeline = [string]$TimelineCache[$cacheKey]
        $ruleName = Get-DataRowText -Row $Row -ColumnName "RuleName"
        if ($ruleName -in @("Ingrowth on install plot", "TreeHistory 10 in initial project period", "Ingrowth with prior tree measurement", "Possible missed ingrowth")) {
            $observedValue = Get-DataRowText -Row $Row -ColumnName "ObservedValue"
            if (-not [string]::IsNullOrWhiteSpace($observedValue)) {
                return "$observedValue; $cachedTimeline"
            }
        }
        return $cachedTimeline
    }

    $orderBy = "[PeriodNumber]"
    if (Test-ColumnExists -Columns $columns -Name "MeasurementID") {
        $orderBy += ", [MeasurementID]"
    }

    try {
        $historyRows = Get-DataTable -Connection $Connection -Sql "SELECT [PeriodNumber], [TreeHistory] FROM [TreeMeasurements] WHERE $whereClause ORDER BY $orderBy"
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($historyRow in $historyRows.Rows) {
            $periodText = Get-DataRowText -Row $historyRow -ColumnName "PeriodNumber"
            if ([string]::IsNullOrWhiteSpace($periodText)) { $periodText = "?" }

            $historyText = Get-DataRowText -Row $historyRow -ColumnName "TreeHistory"
            if ([string]::IsNullOrWhiteSpace($historyText)) { $historyText = "(blank)" }

            [void]$parts.Add("P$periodText=$historyText")
        }

        if ($parts.Count -eq 0) { return "" }
        $timeline = "All periods TreeHistory: " + ([string]::Join("; ", [string[]]$parts.ToArray()))
        $TimelineCache[$cacheKey] = $timeline
        $ruleName = Get-DataRowText -Row $Row -ColumnName "RuleName"
        if ($ruleName -in @("Ingrowth on install plot", "TreeHistory 10 in initial project period", "Ingrowth with prior tree measurement", "Possible missed ingrowth")) {
            $observedValue = Get-DataRowText -Row $Row -ColumnName "ObservedValue"
            if (-not [string]::IsNullOrWhiteSpace($observedValue)) {
                return "$observedValue; $timeline"
            }
        }
        return $timeline
    }
    catch {
        return ""
    }
}

function ConvertTo-XmlText {
    param([object]$Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return "" }
    $text = [string]$Value
    $text = [regex]::Replace($text, "[\x00-\x08\x0B\x0C\x0E-\x1F]", " ")
    if ($text.Length -gt 32000) {
        $text = $text.Substring(0, 31997) + "..."
    }
    return [System.Security.SecurityElement]::Escape($text)
}

function ConvertTo-ExcelColumnName {
    param([int]$Index)
    $name = ""
    while ($Index -gt 0) {
        $mod = ($Index - 1) % 26
        $name = ([char](65 + $mod)) + $name
        $Index = [math]::Floor(($Index - $mod) / 26)
    }
    return $name
}

function Get-SafeWorksheetName {
    param(
        [string]$Name,
        [hashtable]$UsedNames
    )

    $baseName = if ([string]::IsNullOrWhiteSpace($Name)) { "Unspecified" } else { $Name.Trim() }
    $baseName = $baseName -replace "[\\\/\?\*\[\]\:]", " "
    $baseName = $baseName -replace "\s+", " "
    if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = "Unspecified" }
    if ($baseName.Length -gt 31) { $baseName = $baseName.Substring(0, 31) }

    $candidate = $baseName
    $counter = 2
    while ($UsedNames.ContainsKey($candidate.ToLowerInvariant())) {
        $suffix = " $counter"
        $maxBase = 31 - $suffix.Length
        $candidate = $baseName
        if ($candidate.Length -gt $maxBase) { $candidate = $candidate.Substring(0, $maxBase) }
        $candidate += $suffix
        $counter++
    }

    $UsedNames[$candidate.ToLowerInvariant()] = $true
    return $candidate
}

function Get-RowValue {
    param(
        [object]$Row,
        [string]$Name
    )

    if ($Row -is [System.Data.DataRow]) {
        if ($Row.Table.Columns.Contains($Name)) { return $Row[$Name] }
        return ""
    }

    $property = $Row.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return ""
}

function Get-ExportWorksheetGroupName {
    param([object]$Row)

    $fieldName = [string](Get-RowValue -Row $Row -Name "FieldName")
    if ([string]::IsNullOrWhiteSpace($fieldName)) { $fieldName = "Unspecified" }

    $tableName = [string](Get-RowValue -Row $Row -Name "TableName")
    if ($tableName -match "^Regen") {
        return "Regen $fieldName"
    }

    return $fieldName
}

function Test-ExportRowIsReviewedClean {
    param([object]$Row)

    $ruleName = [string](Get-RowValue -Row $Row -Name "RuleName")
    return $ruleName.Equals("Reviewed - no data entry errors found", [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-ExportFindingCount {
    param([object[]]$Rows)

    $count = 0
    foreach ($row in $Rows) {
        if (-not (Test-ExportRowIsReviewedClean -Row $row)) {
            $count++
        }
    }

    return $count
}

function Get-ExportRuleNaturalLanguageSummary {
    param(
        [string]$RuleName,
        [string]$TableName,
        [string]$FieldName
    )

    $fieldLabel = if ([string]::IsNullOrWhiteSpace($FieldName)) { "the selected field" } else { $FieldName }
    switch ($RuleName) {
        "Reviewed - no data entry errors found" { return "DADA reviewed $fieldLabel because it is active in AppColumns and found no data-entry errors for this field." }
        "Invalid CFI code" { return "Nonblank $fieldLabel values were checked against the valid CFI codes in the database template setup." }
        "DMUML code check" { return "DMUML values were checked as three-position codes using only allowed digits with a digit total no greater than 6." }
        "Regen IDBH code check" { return "Regen rows with SpeciesCode greater than 0 were checked for a required IDBH class code of 0, 20, or 40." }
        "Required regen minor plot" { return "Regen rows with SpeciesCode greater than 0 were checked for a required MinorPlot." }
        "Regen minor plot check" { return "Regen MinorPlot values were compared with the project-entered minor plots for timber and woodland seedling/sapling classes." }
        "Rule coverage diagnostic" { return "DADA checked whether plot, tree, and regen cleaning had enough setup context to run fully, including required fields, join keys, AppColumns active flags, woodland species codes, and project-specific allowed lists." }
        "Required regen species" { return "Regen rows with RegenMeasKey, IDBH, StemCount, or MinorPlot data were checked for a required SpeciesCode." }
        "Stem count range check" {
            if ($TableName -eq "RegenMeasurements") {
                return "Entered regen StemCount values were checked for zero, negative, or values above the configured maximum; rows with SpeciesCode greater than 0 were also checked for missing StemCount."
            }
            return "$fieldLabel values were checked for missing, zero, negative, or too-large values."
        }
        "Tree stem count check" { return "Tree StemCount was checked when enabled: entered values must be in range, and woodland or non-timber species require StemCount." }
        "Required plot entry" { return "$fieldLabel was checked as a required plot entry when PlotStatus is 1 or 2." }
        "Required plot remark" { return "PlotRemarks was checked as required when PlotStatus is 7." }
        "PlotStatus progression check" { return "PlotStatus was checked across all periods so the first nonblank status is 2, pre-install periods are blank, and post-install periods are 1, 4, 5, 6, or 7 with no blanks." }
        "Plot data on not-measured plot" { return "Plot measurement and plot custom measurement values were checked so non-allowed data is not recorded for plot statuses 4, 5, 6, or 7." }
        "Regen data on not-measured plot" { return "Regen measurement and regen custom measurement values were checked so data is not recorded for plot statuses 4, 5, 6, or 7. The matching plot status code is included in the export." }
        "Required tree species" { return "Every tree measurement row in the selected cleaning period scope was checked for a joined Trees.SpeciesCode." }
        "Required crown ratio" { return "Crown ratio was checked as required for TreeHistory 0, 5, and 10." }
        "Required crown class" { return "Crown class was checked as required for TreeHistory 0, 5, and 10." }
        "Required radial increment" { return "Radial increment was checked as required for timber TreeHistory 5 missed trees and timber TreeHistory 0, 5, or 10 trees with Problem1 or Problem2 code 121." }
        "Tree data on not-measured plot" { return "Tree measurement and tree custom measurement values were checked so data is not recorded for plot statuses 4, 5, 6, or 7." }
        "Missing DBH" { return "Blank tree DBH was checked against the TreeHistory statuses selected in Run options as requiring DBH." }
        "DBH should be blank" { return "Entered DBH was flagged where TreeHistory 1, 4, or 8 should not have DBH recorded." }
        "IDBH range check" {
            if ($TableName -eq "TreeMeasurements") {
                return "Entered tree DBH values were checked for zero, negative, or above the configured maximum."
            }
            return "$fieldLabel values were checked for missing, zero, negative, or above the configured maximum."
        }
        "Height range check" { return "Tree height was checked for configured range limits, selected TreeHistory required-height rules, rare-species 100% height rules, minor-plot height rules, and timber problem-code no-height rules. Blank height is allowed for selected timber problem-code no-height cases, including TreeHistory 10 ingrowth." }
        "TotalHeight subsample review" { return "TotalHeight subsample counts were checked by plot, SpeciesCode, and 2-inch IDBH class using the selected subsample settings." }
        "TotalHeight subsample order review" { return "TotalHeight subsample order was checked by treating TreeNumber as the north-starting clockwise order and confirming the first required eligible TreeNumbers received heights." }
        "IDBH jump check" { return "Remeasured trees were checked for large DBH jumps against the selected current/prior measurement pair." }
        "Height remeasurement check" { return "Remeasured trees were checked for large height jumps and height decreases." }
        "IDBH shrinkage check" { return "Selected current/prior measurements were compared for live timber DBH shrinkage, skipping woodland species entered in Run options." }
        "Height shrinkage check" { return "Earlier/later measurements were compared for height decreases that may need hand review, excluding allowed live-to-dead mortality shrinkage and previous/current pairs already reported by the remeasurement check." }
        "TreeHistory transition check" { return "TreeHistory transitions were reviewed across all periods for inconsistent live, mortality, harvest, thin, missed, and ingrowth changes." }
        "Lazarus tree check" { return "Trees that appeared to return after mortality, harvest, thinning, or old mortality/harvest were reviewed." }
        "Ingrowth on install plot" { return "TreeHistory 10 was checked against PlotStatus; install plots with PlotStatus 2 should use TreeHistory 0 for new trees." }
        "TreeHistory 10 in initial project period" { return "TreeHistory 10 was checked against the first project measurement period; first-period trees should use TreeHistory 0 instead of ingrowth." }
        "Ingrowth with prior tree measurement" { return "TreeHistory 10 rows were checked for earlier nonblank TreeHistory values that would make ingrowth questionable. Earlier blank TreeHistory rows do not count as prior history for this rule." }
        "Possible missed ingrowth" { return "First-recorded live TreeHistory 0 rows on remeasurement plots with PlotStatus 1 were checked as possible missed ingrowth." }
        "TreeClass/TreeHistory conversion mismatch" { return "TreeClass values were compared with TreeHistory using the supplied crosswalk, and clear mismatches were flagged." }
        "TreeClass/TreeHistory conversion review" { return "TreeClass values were compared with TreeHistory using the supplied crosswalk, and conditional cases were flagged for hand review." }
        "TreeStatus/TreeHistory conversion mismatch" { return "TreeStatus values were compared with TreeHistory using the supplied crosswalk, and clear mismatches were flagged." }
        "TreeStatus/TreeHistory conversion review" { return "TreeStatus values were compared with TreeHistory using the supplied crosswalk, and conditional cases were flagged for hand review." }
        "Problem severity mismatch" { return "Problem and severity pairs were checked so a recorded problem has a matching severity and a recorded severity has a matching problem." }
        "New mortality Severity1 check" { return "Severity1 was checked as code 3 when TreeHistory is 2 or 3 and Problem1 is recorded." }
        "Missing tree measurement" { return "Trees with at least one measurement were checked for missing expected TreeMeasurements rows in the selected measurement periods." }
        "Missing plot measurement" { return "Plots with at least one measurement were checked for missing expected PlotMeasurements rows in the selected measurement periods." }
        "Duplicate selected key" { return "Selected key fields were checked for duplicate records." }
        "Workbook count check" { return "Workbook-style count checks compared parent measurement rows with related measurement and custom-measurement rows." }
        "Project user guide check" { return "Project setup items were compared with the project user guide rules." }
    }

    if (-not [string]::IsNullOrWhiteSpace($RuleName)) {
        return "$fieldLabel was reviewed using the $RuleName rule."
    }
    return "$fieldLabel was reviewed using DADA's configured cleaning checks."
}

function Get-ExportWorksheetSummaryText {
    param(
        [string]$WorksheetName,
        [object[]]$Rows,
        [int]$FindingCount = 0
    )

    $items = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($Rows)) {
        $ruleName = [string](Get-RowValue -Row $row -Name "RuleName")
        $tableName = [string](Get-RowValue -Row $row -Name "TableName")
        $fieldName = [string](Get-RowValue -Row $row -Name "FieldName")
        $summary = Get-ExportRuleNaturalLanguageSummary -RuleName $ruleName -TableName $tableName -FieldName $fieldName
        if ([string]::IsNullOrWhiteSpace($summary)) { continue }
        if (-not $items.Contains($summary)) {
            [void]$items.Add($summary)
        }
    }

    if ($items.Count -eq 0) {
        [void]$items.Add("DADA reviewed this worksheet for configured cleaning checks.")
    }

    $statusText = if ($FindingCount -gt 0) {
        "Rows below are records that need cleaning review."
    }
    else {
        "No cleaning findings were found for this worksheet."
    }

    $text = "What DADA checked: " + ([string]::Join(" ", [string[]]$items.ToArray())) + " $statusText Use PeriodScope to filter rows by current period, previous period, other periods, or all/no-period checks."
    if ($text.Length -gt 1800) {
        $text = $text.Substring(0, 1797) + "..."
    }
    return $text
}

function Get-WorksheetReviewTabColor {
    param([bool]$NeedsReview)

    if ($NeedsReview) { return "FFFF6666" }
    return "FF92D050"
}

function Test-ReportRowsNeedReview {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        $status = [string](Get-RowValue -Row $row -Name "Status")
        if ([string]::IsNullOrWhiteSpace($status)) { continue }
        if (-not $status.Equals("OK", [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Get-ReportReviewNeededCount {
    param([object[]]$Rows)

    $count = 0
    foreach ($row in $Rows) {
        $status = [string](Get-RowValue -Row $row -Name "Status")
        if ([string]::IsNullOrWhiteSpace($status)) { continue }
        if (-not $status.Equals("OK", [System.StringComparison]::OrdinalIgnoreCase)) {
            $count++
        }
    }

    return $count
}

function Test-ReportRowNeedsReview {
    param([object]$Row)

    $status = [string](Get-RowValue -Row $Row -Name "Status")
    if ([string]::IsNullOrWhiteSpace($status)) { return $false }
    return (-not $status.Equals("OK", [System.StringComparison]::OrdinalIgnoreCase))
}

function Get-SummaryPeriodNumberFromRow {
    param([object]$Row)

    foreach ($name in @("PeriodNumber", "CurrentPeriodNumber", "LaterPeriodNumber", "PreviousPeriodNumber", "EarlierPeriodNumber")) {
        $value = [string](Get-RowValue -Row $Row -Name $name)
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $number = 0
        if ([int]::TryParse($value.Trim(), [ref]$number)) {
            return $number
        }
    }

    foreach ($name in @("RecordLabel", "RecordedValue", "ObservedValue", "Notes")) {
        $value = [string](Get-RowValue -Row $Row -Name $name)
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $match = [regex]::Match($value, "(?i)(?:^|[\s;/])P([0-9]{1,2})\b")
        if ($match.Success) {
            $number = 0
            if ([int]::TryParse($match.Groups[1].Value, [ref]$number)) {
                return $number
            }
        }
        $match = [regex]::Match($value, "(?i)\bperiod\s+([0-9]+)\b")
        if ($match.Success) {
            $number = 0
            if ([int]::TryParse($match.Groups[1].Value, [ref]$number)) {
                return $number
            }
        }
    }

    return $null
}

function Get-SummaryPeriodErrorCounts {
    param(
        [object[]]$Rows,
        [string]$RowKind,
        [int]$CurrentPeriodNumber,
        [int]$PreviousPeriodNumber
    )

    $currentCount = 0
    $previousCount = 0
    $otherCount = 0

    foreach ($row in @($Rows)) {
        $needsReview = $false
        if ($RowKind -eq "Report") {
            $needsReview = Test-ReportRowNeedsReview -Row $row
        }
        else {
            $needsReview = -not (Test-ExportRowIsReviewedClean -Row $row)
        }
        if (-not $needsReview) { continue }

        $rowPeriod = Get-SummaryPeriodNumberFromRow -Row $row
        if ($null -ne $rowPeriod -and $CurrentPeriodNumber -gt 0 -and [int]$rowPeriod -eq $CurrentPeriodNumber) {
            $currentCount++
        }
        elseif ($null -ne $rowPeriod -and $PreviousPeriodNumber -gt 0 -and [int]$rowPeriod -eq $PreviousPeriodNumber) {
            $previousCount++
        }
        else {
            $otherCount++
        }
    }

    return [pscustomobject]@{
        Current = $currentCount
        Previous = $previousCount
        Other = $otherCount
    }
}

function Get-ReportPeriodScopeLabel {
    param(
        [object]$Row,
        [int]$CurrentPeriodNumber,
        [int]$PreviousPeriodNumber
    )

    $rowPeriod = Get-SummaryPeriodNumberFromRow -Row $Row
    if ($null -ne $rowPeriod) {
        if ($CurrentPeriodNumber -gt 0 -and [int]$rowPeriod -eq $CurrentPeriodNumber) {
            return "M (period $CurrentPeriodNumber)"
        }
        if ($PreviousPeriodNumber -gt 0 -and [int]$rowPeriod -eq $PreviousPeriodNumber) {
            return "M-1 (period $PreviousPeriodNumber)"
        }
        return "Other period $rowPeriod"
    }

    return "All/no period"
}

function Set-RowPeriodScopeLabel {
    param(
        [object]$Row,
        [string]$Value
    )

    if ($Row -is [System.Data.DataRow]) {
        if ($Row.Table.Columns.Contains("PeriodScope")) {
            $Row["PeriodScope"] = $Value
        }
        return
    }

    $property = $Row.PSObject.Properties["PeriodScope"]
    if ($null -ne $property) {
        $property.Value = $Value
    }
    else {
        $Row | Add-Member -NotePropertyName "PeriodScope" -NotePropertyValue $Value
    }
}

function Set-RowPeriodSummaryColumns {
    param(
        [object]$Row,
        [string]$PeriodScopeLabel,
        [int]$CurrentPeriodNumber,
        [int]$PreviousPeriodNumber
    )

    if ($Row -isnot [System.Data.DataRow]) { return }
    foreach ($columnName in @("ReportCurrentPeriodNumber", "ReportCurrentPeriodErrorCount", "ReportPreviousPeriodNumber", "ReportPreviousPeriodErrorCount")) {
        if (-not $Row.Table.Columns.Contains($columnName)) { return }
    }

    $Row["ReportCurrentPeriodNumber"] = if ($CurrentPeriodNumber -gt 0) { [string]$CurrentPeriodNumber } else { "" }
    $Row["ReportPreviousPeriodNumber"] = if ($PreviousPeriodNumber -gt 0) { [string]$PreviousPeriodNumber } else { "" }
    $Row["ReportCurrentPeriodErrorCount"] = "0"
    $Row["ReportPreviousPeriodErrorCount"] = "0"
    if ([string]::IsNullOrWhiteSpace($PeriodScopeLabel)) { return }

    $isFinding = $true
    try { $isFinding = -not (Test-ExportRowIsReviewedClean -Row $Row) } catch { $isFinding = $true }
    if (-not $isFinding) { return }
    if ($PeriodScopeLabel.StartsWith("M (", [System.StringComparison]::OrdinalIgnoreCase)) {
        $Row["ReportCurrentPeriodErrorCount"] = "1"
    }
    elseif ($PeriodScopeLabel.StartsWith("M-1", [System.StringComparison]::OrdinalIgnoreCase)) {
        $Row["ReportPreviousPeriodErrorCount"] = "1"
    }
}

function Set-PeriodScopeLabelsForRows {
    param(
        [object[]]$Rows,
        [int]$CurrentPeriodNumber,
        [int]$PreviousPeriodNumber
    )

    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $label = Get-ReportPeriodScopeLabel -Row $row -CurrentPeriodNumber $CurrentPeriodNumber -PreviousPeriodNumber $PreviousPeriodNumber
        Set-RowPeriodScopeLabel -Row $row -Value $label
        Set-RowPeriodSummaryColumns -Row $row -PeriodScopeLabel $label -CurrentPeriodNumber $CurrentPeriodNumber -PreviousPeriodNumber $PreviousPeriodNumber
    }
}

function Get-SummaryWorksheetSectionName {
    param([object]$Worksheet)

    $tablesText = ""
    $worksheetName = ""
    try { $tablesText = [string]$Worksheet.Tables } catch { $tablesText = "" }
    try { $worksheetName = [string]$Worksheet.Name } catch { $worksheetName = "" }

    $combined = "$tablesText $worksheetName"
    if ($combined -match "(?i)\bRegen") { return "Regen data" }
    if ($combined -match "(?i)\bTree") { return "Tree data" }
    if ($combined -match "(?i)\bPlot") { return "Plot data" }
    return "Other data"
}

function Get-SummaryWorksheetSectionOrder {
    param([string]$SectionName)

    switch ($SectionName) {
        "Plot data" { return 1 }
        "Tree data" { return 2 }
        "Regen data" { return 3 }
        "Other data" { return 4 }
    }

    return 99
}

function Set-ExportDataRowText {
    param(
        [System.Data.DataTable]$Table,
        [System.Data.DataRow]$Row,
        [string]$ColumnName,
        [object]$Value
    )

    if ($null -eq $Table -or $null -eq $Row -or [string]::IsNullOrWhiteSpace($ColumnName)) { return }
    if (-not $Table.Columns.Contains($ColumnName)) { return }
    $Row[$ColumnName] = if ($null -eq $Value) { "" } else { [string]$Value }
}

function Add-ReviewedCleanExportRows {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.DataTable]$ExportRows
    )

    $script:LastReviewedCleanExportRows = 0
    if ($null -eq $ExportRows) { return 0 }

    $existingGroups = @{}
    foreach ($row in $ExportRows.Rows) {
        $groupName = Get-ExportWorksheetGroupName -Row $row
        if (-not [string]::IsNullOrWhiteSpace($groupName)) {
            $existingGroups[$groupName.ToLowerInvariant()] = $true
        }
    }

    $reviewFields = @(Get-ActiveReviewFieldsForExport -Connection $Connection)
    $added = 0
    foreach ($field in $reviewFields) {
        $groupName = Get-ExportWorksheetGroupName -Row $field
        if ([string]::IsNullOrWhiteSpace($groupName)) { continue }
        $groupKey = $groupName.ToLowerInvariant()
        if ($existingGroups.ContainsKey($groupKey)) { continue }

        $fieldName = [string]$field.FieldName
        $note = "Data for $fieldName was reviewed and no data entry errors were found."
        $newRow = $ExportRows.NewRow()
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "TableName" -Value ([string]$field.TableName)
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "RuleName" -Value ([string]$field.RuleName)
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "RecordLabel" -Value "Reviewed clean"
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "FieldName" -Value $fieldName
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "ObservedValue" -Value ""
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "RecordedValue" -Value $note
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "Message" -Value $note
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "VerificationSql" -Value ""
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "AIMessage" -Value ""
        [void]$ExportRows.Rows.Add($newRow)

        $existingGroups[$groupKey] = $true
        $added++
    }

    foreach ($sourceKind in @("TreeClass", "TreeStatus")) {
        $source = Get-TreeHistoryConversionSourceDefinition -Connection $Connection -SourceKind $sourceKind
        if ($null -eq $source) { continue }

        $fieldName = Get-TreeHistoryConversionFieldName -SourceKind $sourceKind
        $groupKey = $fieldName.ToLowerInvariant()
        if ($existingGroups.ContainsKey($groupKey)) { continue }

        $note = "$sourceKind-to-TreeHistory conversion was reviewed using the crosswalk and no data-entry errors or conditional review rows were found."
        $newRow = $ExportRows.NewRow()
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "TableName" -Value "TreeMeasurements"
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "RuleName" -Value "Reviewed - no data entry errors found"
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "RecordLabel" -Value "Reviewed clean"
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "FieldName" -Value $fieldName
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "ObservedValue" -Value ""
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "RecordedValue" -Value $note
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "Message" -Value $note
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "VerificationSql" -Value ""
        Set-ExportDataRowText -Table $ExportRows -Row $newRow -ColumnName "AIMessage" -Value ""
        [void]$ExportRows.Rows.Add($newRow)

        $existingGroups[$groupKey] = $true
        $added++
    }

    $script:LastReviewedCleanExportRows = $added
    return $added
}

function Test-WorksheetHasTreeRows {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        $tableName = [string](Get-RowValue -Row $row -Name "TableName")
        if ($tableName -in @("Trees", "TreeMeasurements", "TreeCustomMeasurements")) { return $true }
        if (-not [string]::IsNullOrWhiteSpace([string](Get-RowValue -Row $row -Name "TreeNumber"))) { return $true }
    }

    return $false
}

function Test-WorksheetHasPlotRows {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        $tableName = [string](Get-RowValue -Row $row -Name "TableName")
        if ($tableName -in @("Plots", "PlotMeasurements", "PlotCustomMeasurements")) { return $true }
        if (-not [string]::IsNullOrWhiteSpace([string](Get-RowValue -Row $row -Name "PlotStatus"))) { return $true }
        if (-not [string]::IsNullOrWhiteSpace([string](Get-RowValue -Row $row -Name "PlotRemarks"))) { return $true }
    }

    return $false
}

function Test-WorksheetHasRegenRows {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        $tableName = [string](Get-RowValue -Row $row -Name "TableName")
        if ($tableName -match "^Regen") { return $true }
        if (-not [string]::IsNullOrWhiteSpace([string](Get-RowValue -Row $row -Name "RegenMeasKey"))) { return $true }
        if (-not [string]::IsNullOrWhiteSpace([string](Get-RowValue -Row $row -Name "RegenRemarks"))) { return $true }
    }

    return $false
}

function Test-WorksheetColumnHasValue {
    param(
        [object[]]$Rows,
        [string]$ColumnName
    )

    foreach ($row in $Rows) {
        $value = [string](Get-RowValue -Row $row -Name $ColumnName)
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $true }
    }

    return $false
}

function Add-WorksheetHeaderIfMissing {
    param(
        [System.Collections.Generic.List[string]]$Headers,
        [string]$Header
    )

    if (-not $Headers.Contains($Header)) {
        [void]$Headers.Add($Header)
    }
}

function Test-ExportWorksheetIsTotalHeight {
    param([object[]]$Rows)

    foreach ($row in @($Rows)) {
        $fieldName = [string](Get-RowValue -Row $row -Name "FieldName")
        if ($fieldName.Equals("TotalHeight", [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Test-ExportWorksheetIsTreeIdbh {
    param([object[]]$Rows)

    foreach ($row in @($Rows)) {
        $fieldName = [string](Get-RowValue -Row $row -Name "FieldName")
        $tableName = [string](Get-RowValue -Row $row -Name "TableName")
        if ($fieldName.Equals("IDBH", [System.StringComparison]::OrdinalIgnoreCase) -and
            -not ($tableName -match "^Regen")) {
            return $true
        }
    }

    return $false
}

function Test-WorksheetUsesSplitPeriodColumns {
    param([string[]]$Headers)

    return (
        (($Headers -contains "PreviousTotalHeight") -and
            ($Headers -contains "CurrentTotalHeight") -and
            ($Headers -contains "HeightChange")) -or
        (($Headers -contains "PreviousIDBH") -and
            ($Headers -contains "CurrentIDBH") -and
            ($Headers -contains "IDBHJump")) -or
        (($Headers -contains "EarlierIDBH") -and
            ($Headers -contains "LaterIDBH") -and
            ($Headers -contains "Shrinkage"))
    )
}

function ConvertTo-ExportHeaderToken {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $parts = @($Text -split "[^A-Za-z0-9]+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -eq 0) { return "" }

    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($part in $parts) {
        if ($part -cmatch "^[A-Z0-9]+$") {
            [void]$tokens.Add($part)
        }
        elseif ($part.Length -eq 1) {
            [void]$tokens.Add($part.ToUpperInvariant())
        }
        else {
            [void]$tokens.Add($part.Substring(0, 1).ToUpperInvariant() + $part.Substring(1))
        }
    }

    return [string]::Join("", [string[]]$tokens.ToArray())
}

function Get-RecordedValueHeaderName {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        $fieldName = [string](Get-RowValue -Row $row -Name "FieldName")
        $token = ConvertTo-ExportHeaderToken -Text $fieldName
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            return "Recorded${token}Value"
        }
    }

    return "RecordedValue"
}

function Resolve-WorksheetValueColumnName {
    param(
        [string]$Header,
        [bool]$UseReportPeriodColumns = $false
    )

    if ($UseReportPeriodColumns) {
        switch ($Header) {
            "CurrentPeriodNumber" { return "ReportCurrentPeriodNumber" }
            "CurrentPeriodErrorCount" { return "ReportCurrentPeriodErrorCount" }
            "PreviousPeriodNumber" { return "ReportPreviousPeriodNumber" }
            "PreviousPeriodErrorCount" { return "ReportPreviousPeriodErrorCount" }
        }
    }

    if ($Header -eq "Verification SQL") {
        return "VerificationSql"
    }

    if ($Header -match "^Recorded[A-Za-z0-9]+Value$") {
        return "RecordedValue"
    }

    return $Header
}

function Test-RecordedValueWorksheetHeader {
    param([string]$Header)

    return (Resolve-WorksheetValueColumnName -Header $Header) -eq "RecordedValue"
}

function Test-SummaryWorksheetNumberHeader {
    param([string]$Header)

    return ($Header -in @(
        "ErrorCount",
        "CurrentPeriodNumber",
        "CurrentPeriodErrorCount",
        "PreviousPeriodNumber",
        "PreviousPeriodErrorCount",
        "OtherOrAllPeriodErrorCount"
    ))
}

function Test-WorksheetNumericCell {
    param(
        [string]$Header,
        [object]$Value
    )

    $textOnlyHeaders = @(
        "TableName",
        "RuleName",
        "RecordLabel",
        "FieldName",
        "VerificationSql",
        "Verification SQL",
        "AIMessage",
        "TreeRemarks",
        "PlotRemarks",
        "RegenRemarks"
    )

    if ($Header -in $textOnlyHeaders) { return $false }
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return $false }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }

    # Keep leading-zero codes as text so code values such as 020 are not displayed as 20.
    if ($text -match "^[+-]?0\d+$") { return $false }

    return -not [string]::IsNullOrWhiteSpace((ConvertTo-WorksheetNumberText -Value $Value))
}

function Test-SummaryWorksheetCountHeader {
    param([string]$Header)

    return ($Header -in @(
        "ErrorCount",
        "CurrentPeriodErrorCount",
        "PreviousPeriodErrorCount",
        "OtherOrAllPeriodErrorCount"
    ))
}

function Get-SummaryWorksheetColumnFillKind {
    param([string]$Header)

    switch ($Header) {
        "ErrorCount" { return "Yellow" }
        "CurrentPeriodNumber" { return "Blue" }
        "CurrentPeriodErrorCount" { return "Blue" }
        "PreviousPeriodNumber" { return "Green" }
        "PreviousPeriodErrorCount" { return "Green" }
        "OtherOrAllPeriodErrorCount" { return "Yellow" }
    }

    return ""
}

function ConvertTo-WorksheetNumberText {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) { return "" }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    $number = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) { return "" }
        return $number.ToString("0.############", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    if ([double]::TryParse($text, [ref]$number)) {
        if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) { return "" }
        return $number.ToString("0.############", [System.Globalization.CultureInfo]::InvariantCulture)
    }

    return ""
}

function Test-WorksheetPositiveNumber {
    param([object]$Value)

    $numberText = ConvertTo-WorksheetNumberText -Value $Value
    if ([string]::IsNullOrWhiteSpace($numberText)) { return $false }

    $number = 0.0
    if (-not [double]::TryParse($numberText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $false
    }

    return ($number -gt 0)
}

function Get-SummaryWorksheetCellStyleId {
    param(
        [string]$Header,
        [object]$Value,
        [switch]$IsHeader,
        [switch]$SectionStart
    )

    $fillKind = Get-SummaryWorksheetColumnFillKind -Header $Header
    $bold = [bool]$IsHeader
    if ((Test-SummaryWorksheetCountHeader -Header $Header) -and (Test-WorksheetPositiveNumber -Value $Value)) {
        $bold = $true
    }

    if ([string]::IsNullOrWhiteSpace($fillKind)) {
        if ($SectionStart -and $bold) { return 10 }
        if ($SectionStart) { return 9 }
        if ($bold) { return 2 }
        return 0
    }

    switch ($fillKind) {
        "Yellow" {
            if ($SectionStart -and $bold) { return 12 }
            if ($SectionStart) { return 11 }
            if ($bold) { return 4 }
            return 3
        }
        "Blue" {
            if ($SectionStart -and $bold) { return 14 }
            if ($SectionStart) { return 13 }
            if ($bold) { return 6 }
            return 5
        }
        "Green" {
            if ($SectionStart -and $bold) { return 16 }
            if ($SectionStart) { return 15 }
            if ($bold) { return 8 }
            return 7
        }
    }

    if ($SectionStart) { return 9 }
    return 0
}

function Get-WorksheetStyleAttribute {
    param([int]$StyleId)

    if ($StyleId -le 0) { return "" }
    return " s=""$StyleId"""
}

function Add-ExportDataColumnIfMissing {
    param(
        [System.Data.DataTable]$Table,
        [string]$ColumnName
    )

    if ($null -eq $Table -or [string]::IsNullOrWhiteSpace($ColumnName)) { return }
    if (-not $Table.Columns.Contains($ColumnName)) {
        [void]$Table.Columns.Add($ColumnName)
    }
}

function Set-ExportRowColumnValue {
    param(
        [System.Data.DataTable]$Table,
        [System.Data.DataRow]$Row,
        [string]$ColumnName,
        [object]$Value
    )

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return }
    Add-ExportDataColumnIfMissing -Table $Table -ColumnName $ColumnName
    $Row[$ColumnName] = $text
}

function Get-SplitHistoryBaseExportHeaders {
    return @(
        "PreviousPeriodNumber",
        "PreviousIDBH",
        "PreviousTotalHeight",
        "CurrentPeriodNumber",
        "CurrentIDBH",
        "IDBHJump",
        "CurrentTotalHeight",
        "HeightChange",
        "EarlierPeriodNumber",
        "EarlierIDBH",
        "EarlierTotalHeight",
        "LaterPeriodNumber",
        "LaterIDBH",
        "LaterTotalHeight",
        "Shrinkage",
        "PreviousTreeHistory",
        "CurrentTreeHistory",
        "EarlierTreeHistory",
        "LaterTreeHistory"
    )
}

function TryGet-ExportDouble {
    param(
        [string]$Text,
        [ref]$Value
    )

    $parsed = 0.0
    $cleanText = ([string]$Text).Trim()
    if ([string]::IsNullOrWhiteSpace($cleanText)) { return $false }
    if ([double]::TryParse($cleanText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        $Value.Value = $parsed
        return $true
    }
    if ([double]::TryParse($cleanText, [ref]$parsed)) {
        $Value.Value = $parsed
        return $true
    }

    return $false
}

function Format-ExportDouble {
    param([double]$Value)

    if ([Math]::Abs($Value - [Math]::Round($Value)) -lt 0.0000001) {
        return ([int64][Math]::Round($Value)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }

    return $Value.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Set-HeightChangeSplitColumn {
    param(
        [System.Data.DataTable]$Table,
        [System.Data.DataRow]$Row,
        [hashtable]$HeightValues
    )

    if ($null -eq $HeightValues) { return }

    $leftText = ""
    $rightText = ""
    if ($HeightValues.ContainsKey("Previous") -and $HeightValues.ContainsKey("Current")) {
        $leftText = [string]$HeightValues["Previous"]
        $rightText = [string]$HeightValues["Current"]
    }
    elseif ($HeightValues.ContainsKey("Earlier") -and $HeightValues.ContainsKey("Later")) {
        $leftText = [string]$HeightValues["Earlier"]
        $rightText = [string]$HeightValues["Later"]
    }

    if ([string]::IsNullOrWhiteSpace($leftText) -or [string]::IsNullOrWhiteSpace($rightText)) { return }

    $left = 0.0
    $right = 0.0
    if (-not (TryGet-ExportDouble -Text $leftText -Value ([ref]$left))) { return }
    if (-not (TryGet-ExportDouble -Text $rightText -Value ([ref]$right))) { return }
    Set-ExportRowColumnValue -Table $Table -Row $Row -ColumnName "HeightChange" -Value (Format-ExportDouble -Value ($right - $left))
}

function Set-IdbhJumpSplitColumn {
    param(
        [System.Data.DataTable]$Table,
        [System.Data.DataRow]$Row,
        [hashtable]$IdbhValues
    )

    if ($null -eq $IdbhValues) { return }
    if (-not ($IdbhValues.ContainsKey("Previous") -and $IdbhValues.ContainsKey("Current"))) { return }

    $previous = 0.0
    $current = 0.0
    if (-not (TryGet-ExportDouble -Text ([string]$IdbhValues["Previous"]) -Value ([ref]$previous))) { return }
    if (-not (TryGet-ExportDouble -Text ([string]$IdbhValues["Current"]) -Value ([ref]$current))) { return }
    Set-ExportRowColumnValue -Table $Table -Row $Row -ColumnName "IDBHJump" -Value (Format-ExportDouble -Value ($current - $previous))
}

function Set-ObservedHistorySplitColumns {
    param(
        [System.Data.DataTable]$Table,
        [System.Data.DataRow]$Row,
        [string]$ObservedValue
    )

    if ([string]::IsNullOrWhiteSpace($ObservedValue)) { return }

    $heightValues = @{}
    $idbhValues = @{}
    foreach ($match in [regex]::Matches($ObservedValue, "(Previous|Current|Earlier|Later)\s+P([^;\s]+)\s+([A-Za-z0-9]+)=([^;]*)")) {
        $prefix = $match.Groups[1].Value
        $period = $match.Groups[2].Value
        $fieldToken = ConvertTo-ExportHeaderToken -Text $match.Groups[3].Value
        $value = $match.Groups[4].Value
        if ($fieldToken -notin @("IDBH", "TotalHeight", "TreeHistory")) { continue }

        Set-ExportRowColumnValue -Table $Table -Row $Row -ColumnName "${prefix}PeriodNumber" -Value $period
        Set-ExportRowColumnValue -Table $Table -Row $Row -ColumnName "${prefix}${fieldToken}" -Value $value
        if ($fieldToken -eq "TotalHeight") {
            $heightValues[$prefix] = $value
        }
        elseif ($fieldToken -eq "IDBH") {
            $idbhValues[$prefix] = $value
        }
    }

    $hasShrinkageValue = [regex]::IsMatch($ObservedValue, "Shrinkage\s*=", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $hasShrinkageValue) {
        $explicitHeightChange = Get-ObservedKeyValue -ObservedValue $ObservedValue -KeyName "HeightChange"
        if (-not [string]::IsNullOrWhiteSpace($explicitHeightChange)) {
            Set-ExportRowColumnValue -Table $Table -Row $Row -ColumnName "HeightChange" -Value $explicitHeightChange
        }
        else {
            Set-HeightChangeSplitColumn -Table $Table -Row $Row -HeightValues $heightValues
        }
    }

    foreach ($match in [regex]::Matches($ObservedValue, "Shrinkage\s*=\s*([^;]+)")) {
        Set-ExportRowColumnValue -Table $Table -Row $Row -ColumnName "Shrinkage" -Value $match.Groups[1].Value
    }

    $explicitIdbhJump = Get-ObservedKeyValue -ObservedValue $ObservedValue -KeyName "IDBHJump"
    if (-not [string]::IsNullOrWhiteSpace($explicitIdbhJump)) {
        Set-ExportRowColumnValue -Table $Table -Row $Row -ColumnName "IDBHJump" -Value $explicitIdbhJump
    }
    elseif (-not $hasShrinkageValue) {
        Set-IdbhJumpSplitColumn -Table $Table -Row $Row -IdbhValues $idbhValues
    }
}

function Set-TreeHistoryTimelineSplitColumns {
    param(
        [System.Data.DataTable]$Table,
        [System.Data.DataRow]$Row,
        [string]$RecordedValue
    )

    if ([string]::IsNullOrWhiteSpace($RecordedValue)) { return }
    if (-not $RecordedValue.StartsWith("All periods TreeHistory:", [System.StringComparison]::OrdinalIgnoreCase)) { return }

    foreach ($match in [regex]::Matches($RecordedValue, "P([^=;\s]+)=([^;]*)")) {
        $periodToken = ConvertTo-ExportHeaderToken -Text $match.Groups[1].Value
        if ([string]::IsNullOrWhiteSpace($periodToken)) { continue }
        Set-ExportRowColumnValue -Table $Table -Row $Row -ColumnName "TreeHistoryP$periodToken" -Value $match.Groups[2].Value
    }
}

function Set-SplitHistoryExportColumns {
    param(
        [System.Data.DataTable]$Table,
        [System.Data.DataRow]$Row,
        [string]$ObservedValue,
        [string]$RecordedValue
    )

    if (-not (Test-SplitHistoryColumnsSelected)) { return }
    Set-ObservedHistorySplitColumns -Table $Table -Row $Row -ObservedValue $ObservedValue
    Set-TreeHistoryTimelineSplitColumns -Table $Table -Row $Row -RecordedValue $RecordedValue
}

function Get-SplitHistoryDynamicWorksheetHeaders {
    param([object[]]$Rows)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($row in $Rows) {
        if ($row -isnot [System.Data.DataRow]) { continue }
        foreach ($column in $row.Table.Columns) {
            $name = [string]$column.ColumnName
            if (-not ($name -match "^TreeHistoryP[A-Za-z0-9]+$")) { continue }
            if ($names.Contains($name)) { continue }
            if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName $name) {
                [void]$names.Add($name)
            }
        }
        break
    }

    return [string[]](@($names.ToArray()) | Sort-Object `
        @{ Expression = { if ($_ -match "^TreeHistoryP(\d+)$") { [int]$matches[1] } else { 999999 } } },
        @{ Expression = { $_ } })
}

function Get-ExportWorksheetHeaders {
    param([object[]]$Rows)

    $headers = New-Object System.Collections.Generic.List[string]
    $hasTreeRows = Test-WorksheetHasTreeRows -Rows $Rows
    $hasPlotRows = Test-WorksheetHasPlotRows -Rows $Rows
    $hasRegenRows = Test-WorksheetHasRegenRows -Rows $Rows
    $recordedValueHeader = Get-RecordedValueHeaderName -Rows $Rows
    $frontHeaders = if (Test-ExportWorksheetIsTotalHeight -Rows $Rows) {
        @("RuleName", "PreviousPeriodNumber", "PreviousTotalHeight", "CurrentPeriodNumber", "CurrentTotalHeight", "HeightChange", $recordedValueHeader, "Verification SQL")
    }
    elseif (Test-ExportWorksheetIsTreeIdbh -Rows $Rows) {
        $idbhHeaders = New-Object System.Collections.Generic.List[string]
        [void]$idbhHeaders.Add("RuleName")
        foreach ($header in @("PreviousPeriodNumber", "PreviousIDBH", "CurrentPeriodNumber", "CurrentIDBH", "IDBHJump")) {
            if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName $header) {
                [void]$idbhHeaders.Add($header)
            }
        }
        foreach ($header in @("EarlierPeriodNumber", "EarlierIDBH", "LaterPeriodNumber", "LaterIDBH", "Shrinkage")) {
            if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName $header) {
                [void]$idbhHeaders.Add($header)
            }
        }
        [void]$idbhHeaders.Add($recordedValueHeader)
        [void]$idbhHeaders.Add("Verification SQL")
        [string[]]$idbhHeaders.ToArray()
    }
    else {
        @("RuleName", "CurrentPeriodNumber", "CurrentPeriodErrorCount", "PreviousPeriodNumber", "PreviousPeriodErrorCount", $recordedValueHeader, "Verification SQL")
    }
    foreach ($header in $frontHeaders) {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header $header
    }

    foreach ($header in @("PlotNumber", "TreeNumber")) {
        if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName $header) {
            Add-WorksheetHeaderIfMissing -Headers $headers -Header $header
        }
    }

    if ($hasTreeRows -or $hasRegenRows -or (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName "SpeciesCode")) {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header "SpeciesCode"
    }
    if ($hasTreeRows) {
        if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName "TreeHistory") {
            Add-WorksheetHeaderIfMissing -Headers $headers -Header "TreeHistory"
        }
    }

    foreach ($header in @("PeriodNumber", "PeriodScope")) {
        if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName $header) {
            Add-WorksheetHeaderIfMissing -Headers $headers -Header $header
        }
    }

    if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName "MinorPlot") {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header "MinorPlot"
    }
    if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName "PlotStatus") {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header "PlotStatus"
    }

    if ($hasTreeRows) {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header "TreeRemarks"
    }
    if ($hasPlotRows) {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header "PlotRemarks"
    }
    if ($hasRegenRows) {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header "RegenRemarks"
    }

    foreach ($header in Get-SplitHistoryBaseExportHeaders) {
        if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName $header) {
            Add-WorksheetHeaderIfMissing -Headers $headers -Header $header
        }
    }

    foreach ($header in Get-SplitHistoryDynamicWorksheetHeaders -Rows $Rows) {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header $header
    }

    foreach ($header in @("RecordLabel", "TableName", "FieldName", "AuditId", "TreeKey", "RegenMeasKey")) {
        if ($header -in @("RecordLabel", "TableName", "FieldName", "AuditId")) {
            Add-WorksheetHeaderIfMissing -Headers $headers -Header $header
        }
        elseif (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName $header) {
            Add-WorksheetHeaderIfMissing -Headers $headers -Header $header
        }
    }

    if (Test-WorksheetColumnHasValue -Rows $Rows -ColumnName "AIMessage") {
        Add-WorksheetHeaderIfMissing -Headers $headers -Header "AIMessage"
    }

    return [string[]]$headers.ToArray()
}

function Test-ExportWorksheetReviewHeader {
    param([string]$Header)

    return ($Header -in @(
        "RuleName",
        "CurrentPeriodNumber",
        "CurrentPeriodErrorCount",
        "Verification SQL"
    ))
}

function Test-ExportWorksheetContextHeader {
    param([string]$Header)

    if ($Header -in @(
        "PreviousPeriodNumber",
        "PreviousPeriodErrorCount",
        "PlotNumber",
        "TreeNumber",
        "SpeciesCode",
        "TreeHistory",
        "PeriodNumber",
        "PeriodScope",
        "MinorPlot",
        "PlotStatus",
        "TreeRemarks",
        "PlotRemarks",
        "RegenRemarks"
    )) {
        return $true
    }

    if ($Header -in (Get-SplitHistoryBaseExportHeaders)) { return $true }
    if ($Header -match "^TreeHistoryP[A-Za-z0-9]+$") { return $true }
    return $false
}

function Test-ExportWorksheetExtraInfoHeader {
    param([string]$Header)

    return ($Header -in @(
        "RecordLabel",
        "TableName",
        "FieldName",
        "AuditId",
        "TreeKey",
        "RegenMeasKey",
        "AIMessage"
    ))
}

function Get-IdbhRuleFormatStyleId {
    param(
        [string]$RuleName,
        [string]$Header
    )

    if ([string]::IsNullOrWhiteSpace($RuleName) -or [string]::IsNullOrWhiteSpace($Header)) { return 0 }

    $jumpHeaders = @("RuleName", "PreviousPeriodNumber", "PreviousIDBH", "CurrentPeriodNumber", "CurrentIDBH", "IDBHJump")
    $shrinkHeaders = @("RuleName", "EarlierPeriodNumber", "EarlierIDBH", "LaterPeriodNumber", "LaterIDBH", "Shrinkage")

    if ($RuleName.Equals("IDBH jump check", [System.StringComparison]::OrdinalIgnoreCase) -and ($Header -in $jumpHeaders)) {
        return 18
    }
    if ($RuleName.Equals("IDBH shrinkage check", [System.StringComparison]::OrdinalIgnoreCase) -and ($Header -in $shrinkHeaders)) {
        return 20
    }

    return 0
}

function Get-ExportWorksheetCellStyleId {
    param(
        [string]$Header,
        [string]$ValueColumnName,
        [string]$RuleName = ""
    )

    if ($ValueColumnName -eq "RecordedValue") { return 1 }
    $idbhRuleStyleId = Get-IdbhRuleFormatStyleId -RuleName $RuleName -Header $Header
    if ($idbhRuleStyleId -gt 0) { return $idbhRuleStyleId }
    if (Test-ExportWorksheetReviewHeader -Header $Header) { return 17 }
    if (Test-ExportWorksheetContextHeader -Header $Header) { return 18 }
    if (Test-ExportWorksheetExtraInfoHeader -Header $Header) { return 19 }
    return 0
}

function New-WorksheetXml {
    param(
        [string[]]$Headers,
        [object[]]$Rows,
        [string]$TabColor = "",
        [string]$SummaryText = "",
        [switch]$IsSummarySheet
    )

    $lastColumn = ConvertTo-ExcelColumnName $Headers.Count
    $hasSummary = -not [string]::IsNullOrWhiteSpace($SummaryText)
    $headerRow = if ($hasSummary) { 2 } else { 1 }
    $firstDataRow = $headerRow + 1
    $lastRow = [Math]::Max($headerRow, $Rows.Count + $headerRow)
    $useReportPeriodColumns = (-not $IsSummarySheet) -and -not (Test-WorksheetUsesSplitPeriodColumns -Headers $Headers)
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$builder.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    if (-not [string]::IsNullOrWhiteSpace($TabColor)) {
        $safeTabColor = ([string]$TabColor).Trim().ToUpperInvariant()
        if ($safeTabColor -match "^[0-9A-F]{8}$") {
            [void]$builder.AppendLine("<sheetPr><tabColor rgb=""$safeTabColor""/></sheetPr>")
        }
    }
    [void]$builder.AppendLine("<sheetViews><sheetView workbookViewId=""0""><pane ySplit=""$headerRow"" topLeftCell=""A$firstDataRow"" activePane=""bottomLeft"" state=""frozen""/></sheetView></sheetViews>")
    [void]$builder.AppendLine('<cols>')
    for ($i = 1; $i -le $Headers.Count; $i++) {
        $valueColumnName = Resolve-WorksheetValueColumnName -Header $Headers[$i - 1] -UseReportPeriodColumns:$useReportPeriodColumns
        $width = if ($IsSummarySheet -and $Headers[$i - 1] -eq "Worksheet") { 28 } elseif ($IsSummarySheet -and $Headers[$i - 1] -in @("Tables", "Rules")) { 46 } elseif ($IsSummarySheet -and (Test-SummaryWorksheetNumberHeader -Header $Headers[$i - 1])) { 20 } elseif ($Headers[$i - 1] -in @("Message", "VerificationSql", "Verification SQL", "AIMessage", "TreeRemarks", "PlotRemarks", "RegenRemarks")) { 70 } elseif ($valueColumnName -eq "RecordedValue") { 44 } else { 22 }
        [void]$builder.AppendLine("<col min=""$i"" max=""$i"" width=""$width"" customWidth=""1""/>")
    }
    [void]$builder.AppendLine('</cols>')
    [void]$builder.AppendLine('<sheetData>')
    if ($hasSummary) {
        [void]$builder.Append('<row r="1">')
        $summaryAStyle = if ($IsSummarySheet) { ' s="2"' } else { "" }
        $summaryBStyle = if ($IsSummarySheet) { ' s="4"' } else { "" }
        [void]$builder.Append("<c r=""A1""$summaryAStyle t=""inlineStr""><is><t>Checked</t></is></c>")
        [void]$builder.Append("<c r=""B1""$summaryBStyle t=""inlineStr""><is><t>")
        [void]$builder.Append((ConvertTo-XmlText $SummaryText))
        [void]$builder.Append('</t></is></c>')
        [void]$builder.AppendLine('</row>')
    }

    [void]$builder.Append("<row r=""$headerRow"">")
    for ($columnIndex = 1; $columnIndex -le $Headers.Count; $columnIndex++) {
        $cell = (ConvertTo-ExcelColumnName $columnIndex) + $headerRow
        $header = $Headers[$columnIndex - 1]
        $valueColumnName = Resolve-WorksheetValueColumnName -Header $header -UseReportPeriodColumns:$useReportPeriodColumns
        $styleId = if ($IsSummarySheet) { Get-SummaryWorksheetCellStyleId -Header $header -Value "" -IsHeader } else { Get-ExportWorksheetCellStyleId -Header $header -ValueColumnName $valueColumnName }
        $style = Get-WorksheetStyleAttribute -StyleId $styleId
        [void]$builder.Append("<c r=""$cell""$style t=""inlineStr""><is><t>")
        [void]$builder.Append((ConvertTo-XmlText $header))
        [void]$builder.Append('</t></is></c>')
    }
    [void]$builder.AppendLine('</row>')

    $rowNumber = $firstDataRow
    foreach ($row in $Rows) {
        $sectionStart = $false
        if ($IsSummarySheet) {
            try {
                $sectionStart = [bool](Get-RowValue -Row $row -Name "SummarySectionStart")
            }
            catch {
                $sectionStart = $false
            }
        }
        [void]$builder.Append("<row r=""$rowNumber"">")
        $rowRuleName = [string](Get-RowValue -Row $row -Name "RuleName")
        for ($columnIndex = 1; $columnIndex -le $Headers.Count; $columnIndex++) {
            $cell = (ConvertTo-ExcelColumnName $columnIndex) + $rowNumber
            $header = $Headers[$columnIndex - 1]
            $valueColumnName = Resolve-WorksheetValueColumnName -Header $header -UseReportPeriodColumns:$useReportPeriodColumns
            $value = Get-RowValue -Row $row -Name $valueColumnName
            $styleId = if ($IsSummarySheet) { Get-SummaryWorksheetCellStyleId -Header $header -Value $value -SectionStart:($sectionStart) } else { Get-ExportWorksheetCellStyleId -Header $header -ValueColumnName $valueColumnName -RuleName $rowRuleName }
            $style = Get-WorksheetStyleAttribute -StyleId $styleId
            if ($IsSummarySheet -and (Test-SummaryWorksheetNumberHeader -Header $header)) {
                $numberText = ConvertTo-WorksheetNumberText -Value $value
                if ([string]::IsNullOrWhiteSpace($numberText)) {
                    [void]$builder.Append("<c r=""$cell""$style/>")
                }
                else {
                    [void]$builder.Append("<c r=""$cell""$style><v>$numberText</v></c>")
                }
            }
            elseif (Test-WorksheetNumericCell -Header $header -Value $value) {
                $numberText = ConvertTo-WorksheetNumberText -Value $value
                [void]$builder.Append("<c r=""$cell""$style><v>$numberText</v></c>")
            }
            else {
                [void]$builder.Append("<c r=""$cell""$style t=""inlineStr""><is><t>")
                [void]$builder.Append((ConvertTo-XmlText $value))
                [void]$builder.Append('</t></is></c>')
            }
        }
        [void]$builder.AppendLine('</row>')
        $rowNumber++
    }

    [void]$builder.AppendLine('</sheetData>')
    [void]$builder.AppendLine("<autoFilter ref=""A${headerRow}:$lastColumn$lastRow""/>")
    [void]$builder.AppendLine('</worksheet>')
    return $builder.ToString()
}

function New-WorkbookStylesXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="10"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF99"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFCC"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF65C6E8"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFB7E1A1"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top style="thick"><color rgb="FF000000"/></top><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="21"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="5" borderId="0" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0" applyFill="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="8" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="9" borderId="0" xfId="0" applyFill="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
'@
}

function Get-DataRowText {
    param(
        [System.Data.DataRow]$Row,
        [string]$ColumnName
    )

    if ($null -eq $Row -or -not $Row.Table.Columns.Contains($ColumnName)) { return "" }
    $value = $Row[$ColumnName]
    if ($null -eq $value -or $value -is [System.DBNull]) { return "" }
    return [string]$value
}

function Set-IdentityValueIfBlank {
    param(
        [hashtable]$Identity,
        [string]$Name,
        [object]$Value
    )

    if (-not $Identity.ContainsKey($Name)) { $Identity[$Name] = "" }
    if (-not [string]::IsNullOrWhiteSpace([string]$Identity[$Name])) { return }
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return }
    $Identity[$Name] = [string]$Value
}

function Get-FirstDataRow {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$Sql
    )

    try {
        $table = Get-DataTable -Connection $Connection -Sql $Sql
        if ($table.Rows.Count -gt 0) { return $table.Rows[0] }
    }
    catch {
    }

    return $null
}

function Get-RemarkColumnNames {
    param([object[]]$Columns)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($column in $Columns) {
        $name = [string]$column.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $normalized = $name.ToLowerInvariant() -replace "[^a-z0-9]", ""
        if ($normalized -match "(remark|remarks|note|notes|comment|comments)") {
            Add-UniqueCode -Codes $names -Code $name
        }
    }

    return [string[]]$names.ToArray()
}

function Get-FirstMatchingColumnName {
    param(
        [object[]]$Columns,
        [string[]]$CandidateNames
    )

    foreach ($candidate in $CandidateNames) {
        $candidateNormalized = Normalize-FieldName $candidate
        foreach ($column in $Columns) {
            $name = [string]$column.Name
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            if ((Normalize-FieldName $name) -eq $candidateNormalized) {
                return $name
            }
        }
    }

    return ""
}

function Get-ContextValueSelectExpression {
    param(
        [string]$OutputName,
        [string]$PrimaryAlias = "",
        [object[]]$PrimaryColumns = @(),
        [string[]]$PrimaryCandidateNames = @(),
        [string]$SecondaryAlias = "",
        [object[]]$SecondaryColumns = @(),
        [string[]]$SecondaryCandidateNames = @()
    )

    if ($SecondaryCandidateNames.Count -eq 0) {
        $SecondaryCandidateNames = $PrimaryCandidateNames
    }

    $primaryExpression = ""
    $primaryColumn = Get-FirstMatchingColumnName -Columns $PrimaryColumns -CandidateNames $PrimaryCandidateNames
    if (-not [string]::IsNullOrWhiteSpace($PrimaryAlias) -and -not [string]::IsNullOrWhiteSpace($primaryColumn)) {
        $primaryExpression = "Trim(($PrimaryAlias.$(Quote-Name $primaryColumn) & ''))"
    }

    $secondaryExpression = ""
    $secondaryColumn = Get-FirstMatchingColumnName -Columns $SecondaryColumns -CandidateNames $SecondaryCandidateNames
    if (-not [string]::IsNullOrWhiteSpace($SecondaryAlias) -and -not [string]::IsNullOrWhiteSpace($secondaryColumn)) {
        $secondaryExpression = "Trim(($SecondaryAlias.$(Quote-Name $secondaryColumn) & ''))"
    }

    $valueExpression = ""
    if (-not [string]::IsNullOrWhiteSpace($primaryExpression) -and -not [string]::IsNullOrWhiteSpace($secondaryExpression)) {
        $valueExpression = "IIf(Len($primaryExpression) > 0, $primaryExpression, $secondaryExpression)"
    }
    elseif (-not [string]::IsNullOrWhiteSpace($primaryExpression)) {
        $valueExpression = $primaryExpression
    }
    elseif (-not [string]::IsNullOrWhiteSpace($secondaryExpression)) {
        $valueExpression = $secondaryExpression
    }
    else {
        $valueExpression = "''"
    }

    return "$valueExpression AS $(Quote-Name $OutputName)"
}

function Get-SpeciesCodeSelectExpression {
    param(
        [string]$MeasurementAlias = "",
        [object[]]$MeasurementColumns = @(),
        [string]$TreeAlias = "",
        [object[]]$TreeColumns = @()
    )

    return Get-ContextValueSelectExpression `
        -OutputName "SpeciesCode" `
        -PrimaryAlias $MeasurementAlias `
        -PrimaryColumns $MeasurementColumns `
        -PrimaryCandidateNames @("SpeciesCode", "Species") `
        -SecondaryAlias $TreeAlias `
        -SecondaryColumns $TreeColumns `
        -SecondaryCandidateNames @("SpeciesCode", "Species")
}

function Get-PlotStatusSelectExpression {
    param(
        [string]$MeasurementAlias = "",
        [object[]]$MeasurementColumns = @(),
        [string]$PlotAlias = "",
        [object[]]$PlotColumns = @()
    )

    return Get-ContextValueSelectExpression `
        -OutputName "PlotStatus" `
        -PrimaryAlias $MeasurementAlias `
        -PrimaryColumns $MeasurementColumns `
        -PrimaryCandidateNames @("PlotStatus", "PlotStatusCode", "PlotStatusID", "Status", "StatusCode", "StatusID") `
        -SecondaryAlias $PlotAlias `
        -SecondaryColumns $PlotColumns `
        -SecondaryCandidateNames @("PlotStatus", "PlotStatusCode", "PlotStatusID", "Status", "StatusCode", "StatusID")
}

function Get-RemarksSelectExpression {
    param(
        [string]$OutputName,
        [object[]]$Sources
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($source in $Sources) {
        $alias = [string]$source.Alias
        $labelPrefix = [string]$source.Label
        $columns = @($source.Columns)
        if ([string]::IsNullOrWhiteSpace($alias)) { continue }
        if ([string]::IsNullOrWhiteSpace($labelPrefix)) { $labelPrefix = $alias }

        foreach ($columnName in @(Get-RemarkColumnNames -Columns $columns)) {
            $fieldExpression = "($alias.$(Quote-Name $columnName) & '')"
            $label = Sql-Text "${labelPrefix}.${columnName}: "
            [void]$parts.Add("IIf(Len(Trim($fieldExpression)) > 0, $label & $fieldExpression & '; ', '')")
        }
    }

    if ($parts.Count -eq 0) { return "'' AS $(Quote-Name $OutputName)" }
    return "Trim(" + ([string]::Join(" & ", [string[]]$parts.ToArray())) + ") AS $(Quote-Name $OutputName)"
}

function Get-MinorPlotColumnName {
    param([object[]]$Columns)

    foreach ($column in $Columns) {
        $name = [string]$column.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $normalized = Normalize-FieldName $name
        if ($normalized -in @("minorplot", "minorplotnumber")) {
            return $name
        }
    }

    return ""
}

function Get-MinorPlotValueExpression {
    param(
        [string]$MeasurementAlias = "",
        [object[]]$MeasurementColumns = @(),
        [string]$TreeAlias = "",
        [object[]]$TreeColumns = @()
    )

    $measurementColumn = Get-MinorPlotColumnName -Columns $MeasurementColumns
    $treeColumn = Get-MinorPlotColumnName -Columns $TreeColumns

    $measurementExpression = ""
    if (-not [string]::IsNullOrWhiteSpace($MeasurementAlias) -and -not [string]::IsNullOrWhiteSpace($measurementColumn)) {
        $measurementExpression = "Trim(($MeasurementAlias.$(Quote-Name $measurementColumn) & ''))"
    }

    $treeExpression = ""
    if (-not [string]::IsNullOrWhiteSpace($TreeAlias) -and -not [string]::IsNullOrWhiteSpace($treeColumn)) {
        $treeExpression = "Trim(($TreeAlias.$(Quote-Name $treeColumn) & ''))"
    }

    if (-not [string]::IsNullOrWhiteSpace($measurementExpression) -and -not [string]::IsNullOrWhiteSpace($treeExpression)) {
        return "IIf(Len($measurementExpression) > 0, $measurementExpression, $treeExpression)"
    }
    if (-not [string]::IsNullOrWhiteSpace($measurementExpression)) { return $measurementExpression }
    if (-not [string]::IsNullOrWhiteSpace($treeExpression)) { return $treeExpression }
    return ""
}

function Get-MinorPlotSelectExpression {
    param(
        [string]$MeasurementAlias = "",
        [object[]]$MeasurementColumns = @(),
        [string]$TreeAlias = "",
        [object[]]$TreeColumns = @()
    )

    $expression = Get-MinorPlotValueExpression `
        -MeasurementAlias $MeasurementAlias `
        -MeasurementColumns $MeasurementColumns `
        -TreeAlias $TreeAlias `
        -TreeColumns $TreeColumns

    if ([string]::IsNullOrWhiteSpace($expression)) { return "'' AS [MinorPlot]" }
    return "$expression AS [MinorPlot]"
}

function Get-TreeHistorySelectExpression {
    param(
        [string]$Alias,
        [object[]]$Columns
    )

    if (-not [string]::IsNullOrWhiteSpace($Alias) -and (Test-ColumnExists -Columns $Columns -Name "TreeHistory")) {
        return "($Alias.[TreeHistory] & '') AS [TreeHistory]"
    }

    return "'' AS [TreeHistory]"
}

function Get-TreeRemarksSelectExpression {
    param(
        [string]$MeasurementAlias = "",
        [object[]]$MeasurementColumns = @(),
        [string]$TreeAlias = "",
        [object[]]$TreeColumns = @()
    )

    $parts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($MeasurementAlias)) {
        foreach ($columnName in @(Get-RemarkColumnNames -Columns $MeasurementColumns)) {
            $fieldExpression = "($MeasurementAlias.$(Quote-Name $columnName) & '')"
            $label = Sql-Text "TreeMeasurements.${columnName}: "
            [void]$parts.Add("IIf(Len(Trim($fieldExpression)) > 0, $label & $fieldExpression & '; ', '')")
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($TreeAlias)) {
        foreach ($columnName in @(Get-RemarkColumnNames -Columns $TreeColumns)) {
            $fieldExpression = "($TreeAlias.$(Quote-Name $columnName) & '')"
            $label = Sql-Text "Trees.${columnName}: "
            [void]$parts.Add("IIf(Len(Trim($fieldExpression)) > 0, $label & $fieldExpression & '; ', '')")
        }
    }

    if ($parts.Count -eq 0) { return "'' AS [TreeRemarks]" }
    return "Trim(" + ([string]::Join(" & ", [string[]]$parts.ToArray())) + ") AS [TreeRemarks]"
}

function Get-AuditIdentityForRow {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.DataRow]$Row
    )

    $identity = @{
        SourceRowId = Get-DataRowText -Row $Row -ColumnName "SourceRowId"
        PlotNumber = Get-DataRowText -Row $Row -ColumnName "PlotNumber"
        TreeNumber = Get-DataRowText -Row $Row -ColumnName "TreeNumber"
        MinorPlot = Get-DataRowText -Row $Row -ColumnName "MinorPlot"
        PeriodNumber = Get-DataRowText -Row $Row -ColumnName "PeriodNumber"
        PlotKey = Get-DataRowText -Row $Row -ColumnName "PlotKey"
        TreeKey = Get-DataRowText -Row $Row -ColumnName "TreeKey"
        RegenMeasKey = Get-DataRowText -Row $Row -ColumnName "RegenMeasKey"
        SpeciesCode = Get-DataRowText -Row $Row -ColumnName "SpeciesCode"
        PlotStatus = Get-DataRowText -Row $Row -ColumnName "PlotStatus"
        PlotRemarks = Get-DataRowText -Row $Row -ColumnName "PlotRemarks"
        RegenRemarks = Get-DataRowText -Row $Row -ColumnName "RegenRemarks"
        TreeHistory = Get-DataRowText -Row $Row -ColumnName "TreeHistory"
        TreeRemarks = Get-DataRowText -Row $Row -ColumnName "TreeRemarks"
    }

    $tableName = Get-DataRowText -Row $Row -ColumnName "TableName"
    $recordLabel = Get-DataRowText -Row $Row -ColumnName "RecordLabel"
    $sourceRowId = [string]$identity["SourceRowId"]
    $sourceWhere = $null
    if (-not [string]::IsNullOrWhiteSpace($sourceRowId)) {
        $sourceWhere = Sql-Text $sourceRowId
    }

    $lookup = $null
    if ($sourceWhere) {
        switch ($tableName) {
            "Trees" {
                $lookup = Get-FirstDataRow -Connection $Connection -Sql "SELECT [PlotNumber], [TreeNumber], [PlotKey], [TreeKey] FROM [Trees] WHERE ([TreeID] & '') = $sourceWhere"
            }
            "TreeMeasurements" {
                $lookup = Get-FirstDataRow -Connection $Connection -Sql @"
SELECT tm.[PeriodNumber], tm.[TreeKey], tr.[PlotNumber], tr.[TreeNumber], tr.[PlotKey]
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS tr ON tm.[TreeID] = tr.[TreeID]
WHERE (tm.[MeasurementID] & '') = $sourceWhere
"@
            }
            "TreeCustomMeasurements" {
                $lookup = Get-FirstDataRow -Connection $Connection -Sql @"
SELECT tm.[PeriodNumber], tm.[TreeKey], tr.[PlotNumber], tr.[TreeNumber], tr.[PlotKey]
FROM ([TreeCustomMeasurements] AS c
LEFT JOIN [TreeMeasurements] AS tm ON c.[TreeMeasKey] = tm.[TreeMeasKey])
LEFT JOIN [Trees] AS tr ON tm.[TreeID] = tr.[TreeID]
WHERE (c.[MeasurementID] & '') = $sourceWhere
"@
            }
            "PlotMeasurements" {
                $lookup = Get-FirstDataRow -Connection $Connection -Sql @"
SELECT pm.[PeriodNumber], pm.[PlotKey], pl.[PlotNumber]
FROM [PlotMeasurements] AS pm
LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]
WHERE (pm.[MeasurementID] & '') = $sourceWhere
"@
            }
            "PlotCustomMeasurements" {
                $lookup = Get-FirstDataRow -Connection $Connection -Sql @"
SELECT pm.[PeriodNumber], pm.[PlotKey], pl.[PlotNumber]
FROM ([PlotCustomMeasurements] AS c
LEFT JOIN [PlotMeasurements] AS pm ON c.[PlotMeasKey] = pm.[PlotMeasKey])
LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]
WHERE (c.[MeasurementID] & '') = $sourceWhere
"@
            }
            "RegenMeasurements" {
                $lookup = Get-FirstDataRow -Connection $Connection -Sql @"
SELECT p.[PeriodNumber], r.[PlotKey], r.[PlotNumber], r.[SpeciesCode] AS [SpeciesCode], r.[RegenMeasKey] AS [RegenMeasKey]
FROM [RegenMeasurements] AS r
LEFT JOIN [ProjectMeasurementPeriods] AS p ON r.[ProjectPeriodID] = p.[ProjectPeriodID]
WHERE (r.[MeasurementID] & '') = $sourceWhere
"@
            }
            "RegenCustomMeasurements" {
                $lookup = Get-FirstDataRow -Connection $Connection -Sql @"
SELECT p.[PeriodNumber], r.[PlotKey], r.[PlotNumber], r.[SpeciesCode] AS [SpeciesCode], r.[RegenMeasKey] AS [RegenMeasKey]
FROM ([RegenCustomMeasurements] AS c
LEFT JOIN [RegenMeasurements] AS r ON c.[RegenMeasKey] = r.[RegenMeasKey])
LEFT JOIN [ProjectMeasurementPeriods] AS p ON r.[ProjectPeriodID] = p.[ProjectPeriodID]
WHERE (c.[MeasurementID] & '') = $sourceWhere
"@
            }
        }
    }

    if ($lookup) {
        foreach ($name in @("PlotNumber", "TreeNumber", "MinorPlot", "PeriodNumber", "PlotKey", "TreeKey", "RegenMeasKey", "SpeciesCode", "PlotStatus", "PlotRemarks", "RegenRemarks")) {
            Set-IdentityValueIfBlank -Identity $identity -Name $name -Value (Get-DataRowText -Row $lookup -ColumnName $name)
        }
    }

    if ([string]::IsNullOrWhiteSpace([string]$identity["SpeciesCode"])) {
        Set-IdentityValueIfBlank -Identity $identity -Name "SpeciesCode" -Value (Get-ObservedKeyValue -ObservedValue (Get-DataRowText -Row $Row -ColumnName "ObservedValue") -KeyName "SpeciesCode")
    }

    if ([string]::IsNullOrWhiteSpace([string]$identity["TreeNumber"]) -and $recordLabel -match "^\s*([^/]+)/([^ ]+)\s+P(\d+)") {
        Set-IdentityValueIfBlank -Identity $identity -Name "PlotNumber" -Value $matches[1]
        Set-IdentityValueIfBlank -Identity $identity -Name "TreeNumber" -Value $matches[2]
        Set-IdentityValueIfBlank -Identity $identity -Name "PeriodNumber" -Value $matches[3]
    }
    elseif ([string]::IsNullOrWhiteSpace([string]$identity["PlotNumber"]) -and $recordLabel -match "^\s*([^ ]+)\s+P(\d+)") {
        Set-IdentityValueIfBlank -Identity $identity -Name "PlotNumber" -Value $matches[1]
        Set-IdentityValueIfBlank -Identity $identity -Name "PeriodNumber" -Value $matches[2]
    }

    return $identity
}

function New-AuditIdentity {
    param([string]$SourceRowId = "")

    return @{
        SourceRowId = $SourceRowId
        PlotNumber = ""
        TreeNumber = ""
        MinorPlot = ""
        PeriodNumber = ""
        PlotKey = ""
        TreeKey = ""
        RegenMeasKey = ""
        SpeciesCode = ""
        PlotStatus = ""
        PlotRemarks = ""
        RegenRemarks = ""
        TreeHistory = ""
        TreeRemarks = ""
    }
}

function Get-AuditIdentityCacheKey {
    param(
        [string]$TableName,
        [string]$SourceRowId
    )

    return "$TableName|$SourceRowId"
}

function Add-RecordLabelFallbackToIdentity {
    param(
        [hashtable]$Identity,
        [string]$RecordLabel
    )

    if ([string]::IsNullOrWhiteSpace([string]$Identity["TreeNumber"]) -and $RecordLabel -match "^\s*([^/]+)/([^ ]+)\s+P(\d+)") {
        Set-IdentityValueIfBlank -Identity $Identity -Name "PlotNumber" -Value $matches[1]
        Set-IdentityValueIfBlank -Identity $Identity -Name "TreeNumber" -Value $matches[2]
        Set-IdentityValueIfBlank -Identity $Identity -Name "PeriodNumber" -Value $matches[3]
    }
    elseif ([string]::IsNullOrWhiteSpace([string]$Identity["PlotNumber"]) -and $RecordLabel -match "^\s*([^ ]+)\s+P(\d+)") {
        Set-IdentityValueIfBlank -Identity $Identity -Name "PlotNumber" -Value $matches[1]
        Set-IdentityValueIfBlank -Identity $Identity -Name "PeriodNumber" -Value $matches[2]
    }
}

function Add-AuditIdentityLookupRow {
    param(
        [hashtable]$IdentityLookupMap,
        [string]$TableName,
        [System.Data.DataRow]$Row
    )

    $sourceRowId = Get-DataRowText -Row $Row -ColumnName "RowID"
    if ([string]::IsNullOrWhiteSpace($sourceRowId)) { return }

    $identity = New-AuditIdentity -SourceRowId $sourceRowId
    foreach ($name in @("PlotNumber", "TreeNumber", "MinorPlot", "PeriodNumber", "PlotKey", "TreeKey", "RegenMeasKey", "SpeciesCode", "PlotStatus", "PlotRemarks", "RegenRemarks", "TreeHistory", "TreeRemarks")) {
        Set-IdentityValueIfBlank -Identity $identity -Name $name -Value (Get-DataRowText -Row $Row -ColumnName $name)
    }

    $IdentityLookupMap[(Get-AuditIdentityCacheKey -TableName $TableName -SourceRowId $sourceRowId)] = $identity
}

function Add-LatestTreeMeasurementContextToIdentityLookup {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [hashtable]$IdentityLookupMap,
        [string[]]$SourceIds,
        [object[]]$TreeMeasurementColumns
    )

    if ($SourceIds.Count -eq 0) { return }
    foreach ($requiredColumn in @("TreeID", "PeriodNumber")) {
        if (-not (Test-ColumnExists -Columns $TreeMeasurementColumns -Name $requiredColumn)) { return }
    }

    $treeHistorySelect = Get-TreeHistorySelectExpression -Alias "tm" -Columns $TreeMeasurementColumns
    $treeRemarksSelect = Get-TreeRemarksSelectExpression -MeasurementAlias "tm" -MeasurementColumns $TreeMeasurementColumns
    $minorPlotSelect = Get-MinorPlotSelectExpression -MeasurementAlias "tm" -MeasurementColumns $TreeMeasurementColumns

    foreach ($chunk in @(Split-StringList -Values $SourceIds -ChunkSize 150)) {
        $idList = ([string[]]@($chunk | ForEach-Object { Sql-Text $_ })) -join ", "
        if ([string]::IsNullOrWhiteSpace($idList)) { continue }

        $sql = @"
SELECT (tm.[TreeID] & '') AS [RowID], $minorPlotSelect, $treeHistorySelect, $treeRemarksSelect
FROM [TreeMeasurements] AS tm
INNER JOIN (
    SELECT [TreeID], Max([PeriodNumber]) AS [LatestPeriodNumber]
    FROM [TreeMeasurements]
    WHERE ([TreeID] & '') In ($idList)
    GROUP BY [TreeID]
) AS latestPeriod
ON tm.[TreeID] = latestPeriod.[TreeID] AND tm.[PeriodNumber] = latestPeriod.[LatestPeriodNumber]
"@

        try {
            $rows = Get-DataTable -Connection $Connection -Sql $sql
            foreach ($row in $rows.Rows) {
                $sourceRowId = Get-DataRowText -Row $row -ColumnName "RowID"
                if ([string]::IsNullOrWhiteSpace($sourceRowId)) { continue }
                $key = Get-AuditIdentityCacheKey -TableName "Trees" -SourceRowId $sourceRowId
                if (-not $IdentityLookupMap.ContainsKey($key)) { continue }

                $identity = $IdentityLookupMap[$key]
                Set-IdentityValueIfBlank -Identity $identity -Name "MinorPlot" -Value (Get-DataRowText -Row $row -ColumnName "MinorPlot")
                Set-IdentityValueIfBlank -Identity $identity -Name "TreeHistory" -Value (Get-DataRowText -Row $row -ColumnName "TreeHistory")
                $latestRemarks = Get-DataRowText -Row $row -ColumnName "TreeRemarks"
                if (-not [string]::IsNullOrWhiteSpace($latestRemarks)) {
                    $existingRemarks = [string]$identity["TreeRemarks"]
                    if ([string]::IsNullOrWhiteSpace($existingRemarks)) {
                        $identity["TreeRemarks"] = $latestRemarks
                    }
                    elseif ($existingRemarks -notlike "*$latestRemarks*") {
                        $identity["TreeRemarks"] = "$existingRemarks | $latestRemarks"
                    }
                }
            }
        }
        catch {
        }
    }
}

function Split-StringList {
    param(
        [string[]]$Values,
        [int]$ChunkSize = 150
    )

    $chunks = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $Values.Count; $i += $ChunkSize) {
        $end = [Math]::Min($Values.Count - 1, $i + $ChunkSize - 1)
        [void]$chunks.Add(@($Values[$i..$end]))
    }
    return @($chunks.ToArray())
}

function Add-AuditIdentityLookupForTable {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [hashtable]$IdentityLookupMap,
        [string]$TableName,
        [string[]]$SourceIds
    )

    if ($SourceIds.Count -eq 0) { return }

    $treeColumns = @()
    $treeMeasurementColumns = @()
    $plotColumns = @()
    $plotMeasurementColumns = @()
    $regenColumns = @()
    try { $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees") } catch { $treeColumns = @() }
    try { $treeMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements") } catch { $treeMeasurementColumns = @() }
    try { $plotColumns = @(Get-TableColumns -Connection $Connection -TableName "Plots") } catch { $plotColumns = @() }
    try { $plotMeasurementColumns = @(Get-TableColumns -Connection $Connection -TableName "PlotMeasurements") } catch { $plotMeasurementColumns = @() }
    try { $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements") } catch { $regenColumns = @() }
    $speciesCodeFromTree = Get-SpeciesCodeSelectExpression -TreeAlias "t" -TreeColumns $treeColumns
    $speciesCodeFromMeasurement = Get-SpeciesCodeSelectExpression -MeasurementAlias "tm" -MeasurementColumns $treeMeasurementColumns -TreeAlias "tr" -TreeColumns $treeColumns
    $treeHistoryFromTree = Get-TreeHistorySelectExpression -Alias "t" -Columns $treeColumns
    $treeRemarksFromTree = Get-TreeRemarksSelectExpression -TreeAlias "t" -TreeColumns $treeColumns
    $treeHistoryFromMeasurement = Get-TreeHistorySelectExpression -Alias "tm" -Columns $treeMeasurementColumns
    $treeRemarksFromMeasurement = Get-TreeRemarksSelectExpression -MeasurementAlias "tm" -MeasurementColumns $treeMeasurementColumns -TreeAlias "tr" -TreeColumns $treeColumns
    $minorPlotFromTree = Get-MinorPlotSelectExpression -TreeAlias "t" -TreeColumns $treeColumns
    $minorPlotFromMeasurement = Get-MinorPlotSelectExpression -MeasurementAlias "tm" -MeasurementColumns $treeMeasurementColumns -TreeAlias "tr" -TreeColumns $treeColumns
    $minorPlotFromRegen = Get-MinorPlotSelectExpression -MeasurementAlias "r" -MeasurementColumns $regenColumns
    $speciesCodeFromRegen = if (Test-ColumnExists -Columns $regenColumns -Name "SpeciesCode") { "(r.[SpeciesCode] & '') AS [SpeciesCode]" } else { "'' AS [SpeciesCode]" }
    $plotNumberFromPlot = if (Test-ColumnExists -Columns $plotColumns -Name "PlotNumber") { "pl.[PlotNumber]" } else { "'' AS [PlotNumber]" }
    $plotKeyFromPlot = if (Test-ColumnExists -Columns $plotColumns -Name "PlotKey") { "pl.[PlotKey]" } else { "'' AS [PlotKey]" }
    $plotStatusFromPlot = Get-PlotStatusSelectExpression -PlotAlias "pl" -PlotColumns $plotColumns
    $plotRemarksFromPlot = Get-RemarksSelectExpression -OutputName "PlotRemarks" -Sources @(
        @{ Alias = "pl"; Columns = $plotColumns; Label = "Plots" }
    )
    $plotStatusFromMeasurement = Get-PlotStatusSelectExpression -MeasurementAlias "pm" -MeasurementColumns $plotMeasurementColumns -PlotAlias "pl" -PlotColumns $plotColumns
    $plotRemarksFromMeasurement = Get-RemarksSelectExpression -OutputName "PlotRemarks" -Sources @(
        @{ Alias = "pm"; Columns = $plotMeasurementColumns; Label = "PlotMeasurements" },
        @{ Alias = "pl"; Columns = $plotColumns; Label = "Plots" }
    )
    $regenRemarksFromMeasurement = Get-RemarksSelectExpression -OutputName "RegenRemarks" -Sources @(
        @{ Alias = "r"; Columns = $regenColumns; Label = "RegenMeasurements" }
    )

    foreach ($chunk in @(Split-StringList -Values $SourceIds -ChunkSize 150)) {
        $idList = ([string[]]@($chunk | ForEach-Object { Sql-Text $_ })) -join ", "
        if ([string]::IsNullOrWhiteSpace($idList)) { continue }

        $sql = $null
        switch ($TableName) {
            "Trees" {
                $sql = @"
SELECT (t.[TreeID] & '') AS [RowID], t.[PlotNumber], t.[TreeNumber], $speciesCodeFromTree, $minorPlotFromTree, t.[PlotKey], t.[TreeKey], $treeHistoryFromTree, $treeRemarksFromTree
FROM [Trees] AS t
WHERE (t.[TreeID] & '') In ($idList)
"@
            }
            "TreeMeasurements" {
                $sql = @"
SELECT (tm.[MeasurementID] & '') AS [RowID], tm.[PeriodNumber], tm.[TreeKey], tr.[PlotNumber], tr.[TreeNumber], $speciesCodeFromMeasurement, $minorPlotFromMeasurement, tr.[PlotKey], $treeHistoryFromMeasurement, $treeRemarksFromMeasurement
FROM [TreeMeasurements] AS tm
LEFT JOIN [Trees] AS tr ON tm.[TreeID] = tr.[TreeID]
WHERE (tm.[MeasurementID] & '') In ($idList)
"@
            }
            "TreeCustomMeasurements" {
                $sql = @"
SELECT (c.[MeasurementID] & '') AS [RowID], tm.[PeriodNumber], tm.[TreeKey], tr.[PlotNumber], tr.[TreeNumber], $speciesCodeFromMeasurement, $minorPlotFromMeasurement, tr.[PlotKey], $treeHistoryFromMeasurement, $treeRemarksFromMeasurement
FROM ([TreeCustomMeasurements] AS c
LEFT JOIN [TreeMeasurements] AS tm ON c.[TreeMeasKey] = tm.[TreeMeasKey])
LEFT JOIN [Trees] AS tr ON tm.[TreeID] = tr.[TreeID]
WHERE (c.[MeasurementID] & '') In ($idList)
"@
            }
            "Plots" {
                $sql = @"
SELECT (pl.[PlotID] & '') AS [RowID], $plotNumberFromPlot, $plotKeyFromPlot, $plotStatusFromPlot, $plotRemarksFromPlot
FROM [Plots] AS pl
WHERE (pl.[PlotID] & '') In ($idList)
"@
            }
            "PlotMeasurements" {
                $sql = @"
SELECT (pm.[MeasurementID] & '') AS [RowID], pm.[PeriodNumber], pm.[PlotKey], pl.[PlotNumber], $plotStatusFromMeasurement, $plotRemarksFromMeasurement
FROM [PlotMeasurements] AS pm
LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]
WHERE (pm.[MeasurementID] & '') In ($idList)
"@
            }
            "PlotCustomMeasurements" {
                $sql = @"
SELECT (c.[MeasurementID] & '') AS [RowID], pm.[PeriodNumber], pm.[PlotKey], pl.[PlotNumber], $plotStatusFromMeasurement, $plotRemarksFromMeasurement
FROM ([PlotCustomMeasurements] AS c
LEFT JOIN [PlotMeasurements] AS pm ON c.[PlotMeasKey] = pm.[PlotMeasKey])
LEFT JOIN [Plots] AS pl ON pm.[PlotID] = pl.[PlotID]
WHERE (c.[MeasurementID] & '') In ($idList)
"@
            }
            "RegenMeasurements" {
                $sql = @"
SELECT (r.[MeasurementID] & '') AS [RowID], p.[PeriodNumber], r.[PlotKey], r.[PlotNumber], $speciesCodeFromRegen, $minorPlotFromRegen, r.[RegenMeasKey] AS [RegenMeasKey], $regenRemarksFromMeasurement
FROM [RegenMeasurements] AS r
LEFT JOIN [ProjectMeasurementPeriods] AS p ON r.[ProjectPeriodID] = p.[ProjectPeriodID]
WHERE (r.[MeasurementID] & '') In ($idList)
"@
            }
            "RegenCustomMeasurements" {
                $sql = @"
SELECT (c.[MeasurementID] & '') AS [RowID], p.[PeriodNumber], r.[PlotKey], r.[PlotNumber], $speciesCodeFromRegen, $minorPlotFromRegen, r.[RegenMeasKey] AS [RegenMeasKey], $regenRemarksFromMeasurement
FROM ([RegenCustomMeasurements] AS c
LEFT JOIN [RegenMeasurements] AS r ON c.[RegenMeasKey] = r.[RegenMeasKey])
LEFT JOIN [ProjectMeasurementPeriods] AS p ON r.[ProjectPeriodID] = p.[ProjectPeriodID]
WHERE (c.[MeasurementID] & '') In ($idList)
"@
            }
        }

        if ([string]::IsNullOrWhiteSpace($sql)) { continue }

        try {
            $lookupRows = Get-DataTable -Connection $Connection -Sql $sql
            foreach ($row in $lookupRows.Rows) {
                Add-AuditIdentityLookupRow -IdentityLookupMap $IdentityLookupMap -TableName $TableName -Row $row
            }
            if ($TableName -eq "Trees") {
                Add-LatestTreeMeasurementContextToIdentityLookup `
                    -Connection $Connection `
                    -IdentityLookupMap $IdentityLookupMap `
                    -SourceIds ([string[]]$chunk) `
                    -TreeMeasurementColumns $treeMeasurementColumns
            }
        }
        catch {
        }
    }
}

function Get-AuditIdentityLookupMap {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.DataTable]$AuditRows
    )

    $idsByTable = @{}
    foreach ($row in $AuditRows.Rows) {
        $tableName = Get-DataRowText -Row $row -ColumnName "TableName"
        $sourceRowId = Get-DataRowText -Row $row -ColumnName "SourceRowId"
        if ([string]::IsNullOrWhiteSpace($tableName) -or [string]::IsNullOrWhiteSpace($sourceRowId)) { continue }
        if (-not $idsByTable.ContainsKey($tableName)) {
            $idsByTable[$tableName] = New-Object System.Collections.Generic.List[string]
        }
        Add-UniqueCode -Codes $idsByTable[$tableName] -Code $sourceRowId
    }

    $identityLookupMap = @{}
    foreach ($tableName in @($idsByTable.Keys)) {
        Add-AuditIdentityLookupForTable `
            -Connection $Connection `
            -IdentityLookupMap $identityLookupMap `
            -TableName $tableName `
            -SourceIds ([string[]]$idsByTable[$tableName].ToArray())
    }

    return $identityLookupMap
}

function Merge-AuditIdentityFromLookup {
    param(
        [hashtable]$Identity,
        [hashtable]$LookupIdentity
    )

    if ($null -eq $LookupIdentity) { return }
    foreach ($name in @("SourceRowId", "PlotNumber", "TreeNumber", "SpeciesCode", "MinorPlot", "PeriodNumber", "PlotKey", "TreeKey", "RegenMeasKey", "PlotStatus", "PlotRemarks", "RegenRemarks", "TreeHistory", "TreeRemarks")) {
        if ($LookupIdentity.ContainsKey($name)) {
            Set-IdentityValueIfBlank -Identity $Identity -Name $name -Value $LookupIdentity[$name]
        }
    }
}

function Get-AuditIdentityForRowFast {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Data.DataRow]$Row,
        [hashtable]$IdentityLookupMap = $null
    )

    $identity = New-AuditIdentity -SourceRowId (Get-DataRowText -Row $Row -ColumnName "SourceRowId")
    foreach ($name in @("PlotNumber", "TreeNumber", "SpeciesCode", "MinorPlot", "PeriodNumber", "PlotKey", "TreeKey", "RegenMeasKey", "PlotStatus", "PlotRemarks", "RegenRemarks", "TreeHistory", "TreeRemarks")) {
        Set-IdentityValueIfBlank -Identity $identity -Name $name -Value (Get-DataRowText -Row $Row -ColumnName $name)
    }

    $tableName = Get-DataRowText -Row $Row -ColumnName "TableName"
    $sourceRowId = [string]$identity["SourceRowId"]
    if ($null -ne $IdentityLookupMap -and -not [string]::IsNullOrWhiteSpace($tableName) -and -not [string]::IsNullOrWhiteSpace($sourceRowId)) {
        $cacheKey = Get-AuditIdentityCacheKey -TableName $tableName -SourceRowId $sourceRowId
        if ($IdentityLookupMap.ContainsKey($cacheKey)) {
            Merge-AuditIdentityFromLookup -Identity $identity -LookupIdentity $IdentityLookupMap[$cacheKey]
        }
    }

    if ([string]::IsNullOrWhiteSpace([string]$identity["SpeciesCode"])) {
        Set-IdentityValueIfBlank -Identity $identity -Name "SpeciesCode" -Value (Get-ObservedKeyValue -ObservedValue (Get-DataRowText -Row $Row -ColumnName "ObservedValue") -KeyName "SpeciesCode")
    }

    Add-RecordLabelFallbackToIdentity -Identity $identity -RecordLabel (Get-DataRowText -Row $Row -ColumnName "RecordLabel")
    return $identity
}

function Test-AiHelperIsOn {
    if ($null -ne $script:GuideRunSettings) {
        return [bool]$script:GuideRunSettings.AiHelperOn
    }

    try {
        if ($null -ne $aiEnabledCheck -and $aiEnabledCheck.Checked) { return $true }
    }
    catch {
    }

    try {
        if ($null -ne $topAiEnabledCheck -and $topAiEnabledCheck.Checked) { return $true }
    }
    catch {
    }

    return $false
}

function Test-AiProjectGuidanceSelected {
    if ($null -ne $script:GuideRunSettings) {
        return [bool]$script:GuideRunSettings.UseAiProjectGuidance
    }

    try {
        return ($null -ne $aiExportGuidanceCheck -and $aiExportGuidanceCheck.Checked)
    }
    catch {
        return $false
    }
}

function Get-AiRunReadinessIssues {
    $issues = New-Object System.Collections.Generic.List[string]

    if (-not (Test-AiHelperIsOn)) {
        [void]$issues.Add("AI helper is off.")
    }

    if (-not (Test-AiProjectGuidanceSelected)) {
        [void]$issues.Add("Use AI guidance in export is unchecked.")
    }

    try {
        [void](Get-ValidatedAiEndpoint -Endpoint $script:AiDefaultEndpoint)
    }
    catch {
        [void]$issues.Add("The AI endpoint is missing or not valid.")
    }

    if ([string]::IsNullOrWhiteSpace([string]$script:AiDefaultModel)) {
        [void]$issues.Add("The AI model/deployment name is blank.")
    }

    if ([string]::IsNullOrWhiteSpace([string]$script:AiApiKey)) {
        [void]$issues.Add("The AI API key has not been entered.")
    }

    return [string[]]$issues.ToArray()
}

function Confirm-RunWithoutAiIfNeeded {
    $issues = @(Get-AiRunReadinessIssues)
    if ($issues.Count -eq 0) { return $false }

    $issueText = "- " + ([string]::Join("`r`n- ", [string[]]$issues))
    $message = "AI help is not fully set up for this run:`r`n`r`n$issueText`r`n`r`nIf you continue, DADA will run without AI guidance. The workbook will still include built-in cleaning findings, but AIMessage will be blank and AI guidance will be off for this run.`r`n`r`nContinue without AI?"
    $choice = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "Run without AI?",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        Add-Log "Run stopped so AI helper info can be entered before export."
        return $null
    }

    Add-Log "Run continuing without AI guidance. Missing AI setup: $([string]::Join('; ', [string[]]$issues))"
    return $true
}

function Test-AiCompactExportSelected {
    if ($null -ne $script:GuideRunSettings -and $script:GuideRunSettings.PSObject.Properties["AiCompactExport"]) {
        return [bool]$script:GuideRunSettings.AiCompactExport
    }

    try {
        return ($null -eq $aiCompactExportCheck -or $aiCompactExportCheck.Checked)
    }
    catch {
        return $true
    }
}

function Test-SplitHistoryColumnsSelected {
    return $true
}

function Get-ValidatedAiEndpoint {
    param([string]$Endpoint)

    $endpointText = ([string]$Endpoint).Trim().Trim('"').Trim("'")
    if ([string]::IsNullOrWhiteSpace($endpointText)) {
        throw "Enter an AI API endpoint. Default Azure Foundry endpoint: $script:AiDefaultEndpoint."
    }

    if ($endpointText -match "\s") {
        throw "The AI endpoint contains spaces. Paste only the API URL, for example $script:AiDefaultEndpoint."
    }

    if ($endpointText -match "^doi-https[:/\\]" -or $endpointText -match "^https?://doi-https([/:]|$)") {
        throw "The AI endpoint appears to contain a DOI network prefix named 'doi-https'. Remove that prefix and paste the real HTTPS API URL. Default Azure Foundry endpoint: $script:AiDefaultEndpoint."
    }

    if ($endpointText -notmatch "^https://") {
        throw "The AI endpoint must start with https://. Default Azure Foundry endpoint: $script:AiDefaultEndpoint."
    }

    $uri = $null
    if (-not [System.Uri]::TryCreate($endpointText, [System.UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -ne "https" -or
        [string]::IsNullOrWhiteSpace($uri.Host)) {
        throw "The AI endpoint is not a valid URL. Default Azure Foundry endpoint: $script:AiDefaultEndpoint."
    }

    if ($uri.Host -eq "doi-https") {
        throw "The AI endpoint host is 'doi-https', which is not an AI API server. Paste the real HTTPS API URL. Default Azure Foundry endpoint: $script:AiDefaultEndpoint."
    }

    if (-not [string]::IsNullOrWhiteSpace($uri.UserInfo)) {
        throw "Do not put usernames, passwords, or API keys in the AI endpoint URL. Enter the API key only in the API key box."
    }

    if (-not $uri.IsDefaultPort -and $uri.Port -ne 443) {
        throw "The AI endpoint must use the default HTTPS port 443."
    }

    $approvedHosts = @($script:ApprovedAiEndpointHosts | ForEach-Object { ([string]$_).ToLowerInvariant() })
    if (-not ($approvedHosts -contains $uri.Host.ToLowerInvariant())) {
        $approvedText = [string]::Join(", ", [string[]]$approvedHosts)
        throw "This build only allows approved DOI Azure AI endpoint host(s): $approvedText. Ask IT to review and add another host before using a different endpoint."
    }

    return $endpointText
}

function Get-FriendlyAiRequestError {
    param(
        [System.Exception]$Exception,
        [object]$ErrorRecord,
        [string]$Endpoint
    )

    $message = $Exception.Message
    $responseText = ""
    try {
        if ($null -ne $ErrorRecord -and $ErrorRecord.PSObject.Properties["ErrorDetails"] -and
            $null -ne $ErrorRecord.ErrorDetails -and
            -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ErrorDetails.Message)) {
            $responseText = [string]$ErrorRecord.ErrorDetails.Message
        }
    }
    catch {
        $responseText = ""
    }

    try {
        if ([string]::IsNullOrWhiteSpace($responseText) -and $null -ne $Exception.Response) {
            $stream = $Exception.Response.GetResponseStream()
            if ($null -ne $stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                try {
                    $responseText = $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
        }
    }
    catch {
        $responseText = ""
    }

    if (-not [string]::IsNullOrWhiteSpace($responseText)) {
        $responseText = Limit-TextLength -Text ([regex]::Replace($responseText.Trim(), "\s+", " ")) -MaxLength 900
        try {
            $parsedError = $responseText | ConvertFrom-Json
            if ($parsedError.PSObject.Properties["error"]) {
                $errorObject = $parsedError.error
                $parts = New-Object System.Collections.Generic.List[string]
                foreach ($name in @("code", "message", "innererror")) {
                    if ($errorObject.PSObject.Properties[$name] -and -not [string]::IsNullOrWhiteSpace([string]$errorObject.$name)) {
                        [void]$parts.Add("$name=$($errorObject.$name)")
                    }
                }
                if ($parts.Count -gt 0) {
                    $responseText = Limit-TextLength -Text ([string]::Join("; ", [string[]]$parts.ToArray())) -MaxLength 900
                }
            }
        }
        catch {
        }
    }

    if ($message -match "remote name could not be resolved:\s*'([^']+)'") {
        $hostName = $matches[1]
        if ($hostName -eq "doi-https") {
            return "The AI endpoint is pointing to 'doi-https', which is not a real AI API server. Remove the DOI/network prefix and paste the real HTTPS API URL. Default Azure Foundry endpoint: $script:AiDefaultEndpoint."
        }
        return "Could not reach the AI endpoint host '$hostName'. Check the endpoint URL, your network/VPN, and proxy settings."
    }

    if (-not [string]::IsNullOrWhiteSpace($responseText)) {
        return "AI request failed for endpoint '$Endpoint': $message. Service details: $responseText"
    }

    if ($message -match "\(400\)|Bad Request") {
        return "AI request failed for endpoint '$Endpoint': $message. The connection worked, but Azure rejected this specific request. Common causes are a wrong deployment/model name, content filter rejection, or a request that is too large."
    }

    if (-not [string]::IsNullOrWhiteSpace($Endpoint)) {
        return "AI request failed for endpoint '$Endpoint': $message"
    }

    return "AI request failed: $message"
}

function Test-AzureAiEndpoint {
    param([string]$Endpoint)

    $uri = $null
    if (-not [System.Uri]::TryCreate(([string]$Endpoint), [System.UriKind]::Absolute, [ref]$uri)) { return $false }
    if ($uri.Host -match "(^|\.)openai\.azure\.com$") { return $true }
    if ($uri.Host -match "(^|\.)cognitiveservices\.azure\.com$") { return $true }
    if ($uri.Host -match "(^|\.)services\.ai\.azure\.com$") { return $true }
    if ($uri.AbsolutePath -match "/openai/deployments/") { return $true }
    if ($uri.AbsolutePath -match "/models/chat/completions") { return $true }
    return $false
}

function Add-QueryParameterIfMissing {
    param(
        [System.UriBuilder]$Builder,
        [string]$Name,
        [string]$Value
    )

    $query = $Builder.Query
    if ($query.StartsWith("?")) { $query = $query.Substring(1) }
    if ($query -match "(^|&)$([regex]::Escape($Name))=") { return }

    $newPart = "$([System.Uri]::EscapeDataString($Name))=$([System.Uri]::EscapeDataString($Value))"
    if ([string]::IsNullOrWhiteSpace($query)) {
        $Builder.Query = $newPart
    }
    else {
        $Builder.Query = "$query&$newPart"
    }
}

function Get-AiRequestTarget {
    param(
        [string]$Endpoint,
        [string]$Model
    )

    $endpointText = Get-ValidatedAiEndpoint -Endpoint $Endpoint
    $isAzure = Test-AzureAiEndpoint -Endpoint $endpointText
    $endpointUri = [System.Uri]::new($endpointText)
    $isFoundryModelsEndpoint = ($endpointUri.Host -match "(^|\.)services\.ai\.azure\.com$")
    $requestUri = $endpointText
    $useChatCompletions = ($endpointText -match "/chat/completions")
    $includeModelInBody = $false

    if ($isAzure) {
        $useChatCompletions = $true
        $builder = [System.UriBuilder]::new($endpointText)
        $path = $builder.Path
        if ([string]::IsNullOrWhiteSpace($path) -or $path -eq "/") {
            $path = ""
        }
        else {
            $path = $path.TrimEnd("/")
        }

        if ($isFoundryModelsEndpoint -and $path -notmatch "/openai/") {
            if ($path -notmatch "/models/chat/completions$") {
                if ([string]::IsNullOrWhiteSpace($path)) {
                    $path = "/models/chat/completions"
                }
                else {
                    $path = "$path/models/chat/completions"
                }
            }
            $includeModelInBody = $true
            $builder.Path = $path.TrimStart("/")
            Add-QueryParameterIfMissing -Builder $builder -Name "api-version" -Value $script:AzureFoundryModelsApiVersion
            $requestUri = $builder.Uri.AbsoluteUri
        }
        elseif ($path -notmatch "/openai/deployments/") {
            $deployment = [System.Uri]::EscapeDataString($Model)
            $path = "$path/openai/deployments/$deployment/chat/completions"
            $builder.Path = $path.TrimStart("/")
            Add-QueryParameterIfMissing -Builder $builder -Name "api-version" -Value $script:AzureOpenAiApiVersion
            $requestUri = $builder.Uri.AbsoluteUri
        }
        elseif ($path -notmatch "/chat/completions$") {
            $path = "$path/chat/completions"
            $builder.Path = $path.TrimStart("/")
            Add-QueryParameterIfMissing -Builder $builder -Name "api-version" -Value $script:AzureOpenAiApiVersion
            $requestUri = $builder.Uri.AbsoluteUri
        }
        else {
            $builder.Path = $path.TrimStart("/")
            Add-QueryParameterIfMissing -Builder $builder -Name "api-version" -Value $script:AzureOpenAiApiVersion
            $requestUri = $builder.Uri.AbsoluteUri
        }
    }

    return [pscustomobject]@{
        Uri = $requestUri
        IsAzure = $isAzure
        UseChatCompletions = $useChatCompletions
        IncludeModelInBody = $includeModelInBody
    }
}

function New-AiRequestHeaders {
    param(
        [string]$ApiKey,
        [bool]$IsAzure
    )

    if ($IsAzure) {
        return @{ "api-key" = $ApiKey }
    }

    return @{ "Authorization" = "Bearer $ApiKey" }
}

function New-AiChatBody {
    param(
        [string]$Model,
        [string]$SystemText,
        [string]$UserText,
        [bool]$IsAzure,
        [bool]$IncludeModelInBody = $false
    )

    $body = @{
        messages = @(
            @{ role = "system"; content = (ConvertTo-AiSafeText -Value $SystemText) },
            @{ role = "user"; content = (ConvertTo-AiSafeText -Value $UserText) }
        )
    }

    if (-not $IsAzure -or $IncludeModelInBody) {
        $body["model"] = $Model
    }

    return $body
}

function Assert-AiProjectGuidanceReady {
    if (-not (Test-AiHelperIsOn)) {
        throw "Turn on AI helper before using AI guidance in the export."
    }

    $endpoint = $script:AiDefaultEndpoint
    $model = $script:AiDefaultModel
    $apiKey = $script:AiApiKey
    $manualPath = $script:ProjectManualPath
    if ($null -ne $script:GuideRunSettings) {
        $endpoint = [string]$script:GuideRunSettings.AiEndpoint
        $model = [string]$script:GuideRunSettings.AiModel
        $apiKey = [string]$script:GuideRunSettings.AiApiKey
        $manualPath = [string]$script:GuideRunSettings.ProjectManualPath
    }

    [void](Get-ValidatedAiEndpoint -Endpoint $endpoint)
    if ([string]::IsNullOrWhiteSpace($model)) { throw "Open AI helper and enter a model." }
    if ([string]::IsNullOrWhiteSpace($apiKey)) { throw "Open AI helper and enter an API key. The key is only kept while this app window is open." }
    if (-not [string]::IsNullOrWhiteSpace($manualPath)) {
        [void](Get-ProjectManualText)
    }
}

function Invoke-AiExportGuidanceRequest {
    param(
        [string]$Endpoint,
        [string]$ApiKey,
        [string]$Model,
        [string]$InputText,
        [switch]$DisableOptionalChatParameters
    )

    if ([string]::IsNullOrWhiteSpace($ApiKey)) { throw "Enter an API key." }
    if ([string]::IsNullOrWhiteSpace($Model)) { throw "Enter a model." }
    $target = Get-AiRequestTarget -Endpoint $Endpoint -Model $Model

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$instructions = @"
You help prepare a forest inventory data cleaning export.
Use the coded rule/template context as authoritative for allowed values and app checks. Use the project manual excerpt when it is relevant, and use previous/current measurement context when present.
Forestry words such as mortality, harvest, thinning, dead, and live refer only to tree status codes in inventory data.
Return JSON only with two string properties: ai_message and suggested_edit.
ai_message must be one plain sentence, 45 words or fewer, with no markdown, no bullets, and no preamble.
suggested_edit must be only the exact replacement value to put in the field, with no explanatory words or units. Use an empty string when the exact value is not supported by the finding/manual/current/prior data. Use NULL only when the field should be cleared.
Put explanations only in ai_message. Never put an instruction, sentence, or explanation in suggested_edit.
"@

    $headers = New-AiRequestHeaders -ApiKey $ApiKey -IsAzure ([bool]$target.IsAzure)

    if ($target.UseChatCompletions) {
        $body = New-AiChatBody `
            -Model $Model `
            -SystemText $instructions `
            -UserText $InputText `
            -IsAzure:([bool]$target.IsAzure) `
            -IncludeModelInBody:([bool]$target.IncludeModelInBody)
        if (-not $DisableOptionalChatParameters) {
            $body["max_tokens"] = 220
            $body["temperature"] = 0.2
        }
    }
    else {
        $body = @{
            model = $Model
            instructions = (ConvertTo-AiSafeText -Value $instructions)
            input = (ConvertTo-AiSafeText -Value $InputText)
            max_output_tokens = 180
        }
    }

    $json = $body | ConvertTo-Json -Depth 10
    $jsonBody = ConvertTo-Utf8JsonBody -Json $json
    try {
        $response = Invoke-RestMethod -Method Post -Uri ([string]$target.Uri) -Headers $headers -ContentType "application/json; charset=utf-8" -Body $jsonBody -TimeoutSec 90
    }
    catch {
        $friendlyError = Get-FriendlyAiRequestError -Exception $_.Exception -ErrorRecord $_ -Endpoint ([string]$target.Uri)
        if ($target.UseChatCompletions -and -not $DisableOptionalChatParameters -and (Test-AiBadRequestMessage -Message $friendlyError)) {
            Assert-GuideRunNotCanceled
            try {
                return Invoke-AiExportGuidanceRequest `
                    -Endpoint $Endpoint `
                    -ApiKey $ApiKey `
                    -Model $Model `
                    -InputText $InputText `
                    -DisableOptionalChatParameters
            }
            catch {
                throw "AI request was rejected with normal settings, and the retry without optional model settings also failed. Normal request: $friendlyError Retry without optional settings: $($_.Exception.Message)"
            }
        }

        throw $friendlyError
    }
    $text = Get-OpenAiResponseText -Response $response
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw "The model returned an empty response. Use Test connection to confirm the endpoint, model/deployment name, and API key."
    }
    return (Limit-TextLength -Text $text.Trim() -MaxLength 1200)
}

function Invoke-AiExportGuidanceRequestWithRetry {
    param(
        [string]$Endpoint,
        [string]$ApiKey,
        [string]$Model,
        [string]$InputText,
        [string]$CompactInputText,
        [string]$UltraCompactInputText,
        [string]$FinalFallbackInputText,
        [string]$InputLabel = "full",
        [string]$CompactLabel = "compact",
        [string]$UltraCompactLabel = "tiny-manual-digest",
        [string]$FinalFallbackLabel = "tiny-no-manual"
    )

    try {
        $script:LastAiExportGuidanceAttemptLabel = $InputLabel
        return Invoke-AiExportGuidanceRequest `
            -Endpoint $Endpoint `
            -ApiKey $ApiKey `
            -Model $Model `
            -InputText $InputText
    }
    catch {
        $firstError = $_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($CompactInputText) -or -not (Test-AiBadRequestMessage -Message $firstError)) {
            throw
        }

        Assert-GuideRunNotCanceled
        try {
            $script:LastAiExportGuidanceAttemptLabel = $CompactLabel
            $compactResponse = Invoke-AiExportGuidanceRequest `
                -Endpoint $Endpoint `
                -ApiKey $ApiKey `
                -Model $Model `
                -InputText $CompactInputText
            return $compactResponse
        }
        catch {
            $secondError = $_.Exception.Message
            if (-not [string]::IsNullOrWhiteSpace($UltraCompactInputText) -and (Test-AiBadRequestMessage -Message $secondError)) {
                Assert-GuideRunNotCanceled
                try {
                    $script:LastAiExportGuidanceAttemptLabel = $UltraCompactLabel
                    return Invoke-AiExportGuidanceRequest `
                        -Endpoint $Endpoint `
                        -ApiKey $ApiKey `
                        -Model $Model `
                        -InputText $UltraCompactInputText
                }
                catch {
                    $thirdError = $_.Exception.Message
                    if (-not [string]::IsNullOrWhiteSpace($FinalFallbackInputText) -and (Test-AiBadRequestMessage -Message $thirdError)) {
                        Assert-GuideRunNotCanceled
                        try {
                            $script:LastAiExportGuidanceAttemptLabel = $FinalFallbackLabel
                            return Invoke-AiExportGuidanceRequest `
                                -Endpoint $Endpoint `
                                -ApiKey $ApiKey `
                                -Model $Model `
                                -InputText $FinalFallbackInputText
                        }
                        catch {
                            $fourthError = $_.Exception.Message
                            throw "AI export request was rejected after four attempts. First: $firstError Compact: $secondError Tiny manual digest: $thirdError Tiny no-manual: $fourthError"
                        }
                    }

                    throw "AI export request was rejected after three attempts. First: $firstError Compact: $secondError Tiny manual digest: $thirdError"
                }
            }

            throw "AI export request was rejected, and the compact retry also failed. First: $firstError Compact: $secondError"
        }
    }
}

function Get-AiProjectGuidanceForExportFinding {
    param(
        [string]$TableName,
        [string]$FieldName,
        [string]$RuleName,
        [string]$Message,
        [string]$SuggestedEdit,
        [string]$RuleSuggestedEdit,
        [string]$SuggestedValue,
        [string]$SuggestedValueSummary,
        [string]$ObservedValue,
        [string]$RecordLabel,
        [string]$RuleContext,
        [int]$GroupCount = 1,
        [switch]$AllowSuggestedValue
    )

    Assert-GuideRunNotCanceled
    $manualLoadedFrom = $script:ProjectManualLoadedFrom
    $endpoint = $script:AiDefaultEndpoint
    $model = $script:AiDefaultModel
    $apiKey = $script:AiApiKey
    $useCompactExport = Test-AiCompactExportSelected
    if ($null -ne $script:GuideRunSettings) {
        $manualLoadedFrom = [string]$script:GuideRunSettings.ProjectManualLoadedFrom
        $endpoint = [string]$script:GuideRunSettings.AiEndpoint
        $model = [string]$script:GuideRunSettings.AiModel
        $apiKey = [string]$script:GuideRunSettings.AiApiKey
    }

    $cacheParts = New-Object System.Collections.Generic.List[string]
    foreach ($part in @($manualLoadedFrom, $TableName, $FieldName, $RuleName, $Message, $SuggestedEdit, $RuleSuggestedEdit, $RuleContext)) {
        [void]$cacheParts.Add([string]$part)
    }
    if ($AllowSuggestedValue) {
        foreach ($part in @($ObservedValue, $SuggestedValue)) {
            [void]$cacheParts.Add([string]$part)
        }
    }
    $cacheKey = [string]::Join("|", [string[]]$cacheParts.ToArray())

    if ($script:AiGuidanceCache.ContainsKey($cacheKey)) {
        $cached = $script:AiGuidanceCache[$cacheKey]
        if ($cached -is [string]) {
            return [pscustomobject]@{
                AIMessage = [string]$cached
                SuggestedEdit = ""
                Guidance = [string]$cached
                SuggestedValue = ""
                ManualSnippetAvailable = $false
                ManualSnippetIncluded = $false
                PromptAttempt = "cache"
            }
        }
        return $cached
    }

    $manualSnippet = Get-ProjectManualSnippet `
        -TableName $TableName `
        -FieldName $FieldName `
        -RuleName $RuleName `
        -Message $Message `
        -MaxLength 2500
    $compactManualSnippet = Get-ProjectManualSnippet `
        -TableName $TableName `
        -FieldName $FieldName `
        -RuleName $RuleName `
        -Message $Message `
        -MaxLength 500 `
        -Dense
    if ([string]::IsNullOrWhiteSpace($compactManualSnippet) -and -not [string]::IsNullOrWhiteSpace($manualSnippet)) {
        $compactManualSnippet = Limit-TextLength -Text $manualSnippet -MaxLength 500
    }
    $manualDigest = Get-ProjectManualDigestForFinding `
        -TableName $TableName `
        -FieldName $FieldName `
        -RuleName $RuleName `
        -Message $Message `
        -MaxLength 260
    if ([string]::IsNullOrWhiteSpace($manualDigest) -and -not [string]::IsNullOrWhiteSpace($compactManualSnippet)) {
        $manualDigest = Convert-ManualSnippetToDigest -Snippet $compactManualSnippet -MaxLength 260
    }
    $tinyManualDigest = Limit-TextLength -Text $manualDigest -MaxLength 160
    $manualSnippetAvailable = (-not [string]::IsNullOrWhiteSpace($manualDigest))

    if ([string]::IsNullOrWhiteSpace($SuggestedValueSummary)) {
        if ([string]::IsNullOrWhiteSpace($SuggestedValue)) {
            $SuggestedValueSummary = "No value-only suggestion is currently supported for this group."
        }
        else {
            $SuggestedValueSummary = "Current value-only suggestion: $SuggestedValue"
        }
    }

    $suggestedValueInstruction = if ($AllowSuggestedValue) {
        "If one exact replacement value is supported for every row in this group, return it in suggested_edit. Otherwise return an empty string."
    }
    else {
        "This grouped finding may cover rows with different recorded or suggested values. Return an empty string for suggested_edit."
    }

    $inputText = @"
Audit finding group:
Rows represented: $GroupCount
Table: $TableName
Field: $FieldName
Rule: $RuleName
Example record: $RecordLabel
Example recorded value: $(Limit-TextLength -Text $ObservedValue -MaxLength 350)
Original export message: $(Limit-TextLength -Text $Message -MaxLength 500)
Built-in reviewer guidance: $(Limit-TextLength -Text $RuleSuggestedEdit -MaxLength 500)
Rule-based value hint for example row: $(Limit-TextLength -Text $SuggestedEdit -MaxLength 80)
$SuggestedValueSummary
$suggestedValueInstruction

Coded app/template rule context:
$(Limit-TextLength -Text $RuleContext -MaxLength 3000)

Relevant project manual excerpt:
$(Limit-TextLength -Text $manualSnippet -MaxLength 2500)
"@

    $compactInputText = @"
Audit finding group:
Rows represented: $GroupCount
Table: $TableName
Field: $FieldName
Rule: $RuleName
Example record: $RecordLabel
Example recorded value: $(Limit-TextLength -Text $ObservedValue -MaxLength 180)
Original export message: $(Limit-TextLength -Text $Message -MaxLength 240)
Built-in reviewer guidance: $(Limit-TextLength -Text $RuleSuggestedEdit -MaxLength 240)
Rule-based value hint for example row: $(Limit-TextLength -Text $SuggestedEdit -MaxLength 80)
$SuggestedValueSummary
$suggestedValueInstruction

Compact coded rule context:
$(Limit-TextLength -Text $RuleContext -MaxLength 700)

Compact project manual digest:
$(Limit-TextLength -Text $manualDigest -MaxLength 260)
"@

    $tinyManualDigestInputText = @"
Audit finding group:
Rows represented: $GroupCount
Table: $TableName
Field: $FieldName
Rule: $RuleName
Example record: $RecordLabel
Example recorded value: $(Limit-TextLength -Text $ObservedValue -MaxLength 120)
Original export message: $(Limit-TextLength -Text $Message -MaxLength 160)
Built-in reviewer guidance: $(Limit-TextLength -Text $RuleSuggestedEdit -MaxLength 160)
Rule-based value hint for example row: $(Limit-TextLength -Text $SuggestedEdit -MaxLength 80)
$SuggestedValueSummary
$suggestedValueInstruction

Tiny coded rule context:
$(Limit-TextLength -Text $RuleContext -MaxLength 180)

Tiny project manual digest:
$tinyManualDigest
"@

    $noManualInputText = @"
Audit finding group:
Rows represented: $GroupCount
Table: $TableName
Field: $FieldName
Rule: $RuleName
Example record: $RecordLabel
Example recorded value: $(Limit-TextLength -Text $ObservedValue -MaxLength 120)
Original export message: $(Limit-TextLength -Text $Message -MaxLength 160)
Built-in reviewer guidance: $(Limit-TextLength -Text $RuleSuggestedEdit -MaxLength 160)
Rule-based value hint for example row: $(Limit-TextLength -Text $SuggestedEdit -MaxLength 80)
$SuggestedValueSummary
$suggestedValueInstruction

Tiny coded rule context:
$(Limit-TextLength -Text $RuleContext -MaxLength 240)

No project manual excerpt is included in this retry.
"@

    if ($useCompactExport) {
        $script:LastAiExportGuidanceAttemptLabel = ""
        $rawResponse = Invoke-AiExportGuidanceRequestWithRetry `
            -Endpoint $endpoint `
            -ApiKey $apiKey `
            -Model $model `
            -InputText $compactInputText `
            -CompactInputText $tinyManualDigestInputText `
            -UltraCompactInputText $noManualInputText `
            -InputLabel "compact" `
            -CompactLabel "tiny-manual-digest" `
            -UltraCompactLabel "tiny-no-manual"
    }
    else {
        $script:LastAiExportGuidanceAttemptLabel = ""
        $rawResponse = Invoke-AiExportGuidanceRequestWithRetry `
            -Endpoint $endpoint `
            -ApiKey $apiKey `
            -Model $model `
            -InputText $inputText `
            -CompactInputText $compactInputText `
            -UltraCompactInputText $tinyManualDigestInputText `
            -FinalFallbackInputText $noManualInputText `
            -InputLabel "full" `
            -CompactLabel "compact" `
            -UltraCompactLabel "tiny-manual-digest" `
            -FinalFallbackLabel "tiny-no-manual"
    }

    $promptAttempt = [string]$script:LastAiExportGuidanceAttemptLabel
    if ([string]::IsNullOrWhiteSpace($promptAttempt)) {
        $promptAttempt = if ($useCompactExport) { "compact" } else { "full" }
    }
    $manualSnippetIncluded = ($manualSnippetAvailable -and $promptAttempt -ne "tiny-no-manual")

    Assert-GuideRunNotCanceled
    $aiMessageFromAi = ""
    $suggestedEditValueFromAi = ""
    try {
        $jsonText = $rawResponse.Trim()
        if ($jsonText -match '(?s)```(?:json)?\s*(.*?)\s*```') {
            $jsonText = $matches[1].Trim()
        }
        $parsed = $jsonText | ConvertFrom-Json

        if ($parsed.PSObject.Properties["ai_message"]) {
            $aiMessageFromAi = [string]$parsed.ai_message
        }
        elseif ($parsed.PSObject.Properties["guidance"]) {
            $aiMessageFromAi = [string]$parsed.guidance
        }
        elseif ($parsed.PSObject.Properties["message"]) {
            $aiMessageFromAi = [string]$parsed.message
        }

        if ($AllowSuggestedValue -and $parsed.PSObject.Properties["suggested_edit"]) {
            $suggestedEditValueFromAi = Normalize-AiSuggestedEditValue $parsed.suggested_edit
            if ([string]::IsNullOrWhiteSpace($suggestedEditValueFromAi) -and [string]::IsNullOrWhiteSpace($aiMessageFromAi)) {
                $aiMessageFromAi = [string]$parsed.suggested_edit
            }
        }

        if ($AllowSuggestedValue -and [string]::IsNullOrWhiteSpace($suggestedEditValueFromAi) -and $parsed.PSObject.Properties["suggested_value"]) {
            $suggestedEditValueFromAi = Normalize-AiSuggestedEditValue $parsed.suggested_value
        }
    }
    catch {
        $aiMessageFromAi = $rawResponse
    }

    $aiMessageFromAi = [regex]::Replace($aiMessageFromAi.Trim(), "\s+", " ")
    $aiMessageFromAi = $aiMessageFromAi -replace "^(AIProjectGuidance|AIMessage|Guidance|Message|SuggestedEdit|Suggested edit)\s*:\s*", ""
    $aiMessageFromAi = Limit-TextLength -Text $aiMessageFromAi -MaxLength 600
    if ([string]::IsNullOrWhiteSpace($aiMessageFromAi) -and -not [string]::IsNullOrWhiteSpace($suggestedEditValueFromAi)) {
        $aiMessageFromAi = "AI suggested this value from the rule, manual, and measurement context; verify it against the original field record before updating."
    }

    $result = [pscustomobject]@{
        AIMessage = $aiMessageFromAi
        SuggestedEdit = $suggestedEditValueFromAi
        Guidance = $aiMessageFromAi
        SuggestedValue = $suggestedEditValueFromAi
        ManualSnippetAvailable = [bool]$manualSnippetAvailable
        ManualSnippetIncluded = [bool]$manualSnippetIncluded
        PromptAttempt = $promptAttempt
    }
    $script:AiGuidanceCache[$cacheKey] = $result
    return $result
}

function Get-AiProjectGuidanceForAuditRow {
    param(
        [System.Data.DataRow]$Row,
        [string]$SuggestedEdit,
        [string]$SuggestedValue
    )

    $tableName = Get-DataRowText -Row $Row -ColumnName "TableName"
    $fieldName = Get-DataRowText -Row $Row -ColumnName "FieldName"
    $ruleName = Get-DataRowText -Row $Row -ColumnName "RuleName"
    $message = Get-DataRowText -Row $Row -ColumnName "Message"
    $observedValue = Get-DataRowText -Row $Row -ColumnName "ObservedValue"
    $recordLabel = Get-DataRowText -Row $Row -ColumnName "RecordLabel"
    $summary = if ([string]::IsNullOrWhiteSpace($SuggestedValue)) {
        "No value-only suggestion is currently supported for this finding."
    }
    else {
        "Current value-only suggestion: $SuggestedValue"
    }

    return Get-AiProjectGuidanceForExportFinding `
        -TableName $tableName `
        -FieldName $fieldName `
        -RuleName $ruleName `
        -Message $message `
        -SuggestedEdit $SuggestedValue `
        -RuleSuggestedEdit $SuggestedEdit `
        -SuggestedValue $SuggestedValue `
        -SuggestedValueSummary $summary `
        -ObservedValue $observedValue `
        -RecordLabel $recordLabel `
        -RuleContext "" `
        -GroupCount 1 `
        -AllowSuggestedValue
}

function Get-UniqueExportRowValues {
    param(
        [object[]]$Rows,
        [string]$ColumnName,
        [switch]$NonBlankOnly
    )

    $values = New-Object System.Collections.Generic.List[string]
    foreach ($row in $Rows) {
        $value = [string](Get-RowValue -Row $row -Name $ColumnName)
        if ($NonBlankOnly -and [string]::IsNullOrWhiteSpace($value)) { continue }
        if (-not $values.Contains($value)) {
            [void]$values.Add($value)
        }
    }

    return @($values.ToArray())
}

function Get-AiSuggestedValueSummaryForGroup {
    param([object[]]$Rows)

    $values = @(Get-UniqueExportRowValues -Rows $Rows -ColumnName "SuggestedValue" -NonBlankOnly)
    if ($values.Count -eq 0) {
        return "No value-only suggestion is currently supported for this group."
    }
    if ($values.Count -eq 1) {
        return "Current value-only suggestion for this group: $(Limit-TextLength -Text $values[0] -MaxLength 80)"
    }

    $sampleValues = [string]::Join(", ", [string[]]@($values | Select-Object -First 5))
    return "Current value-only suggestions vary by row (examples: $(Limit-TextLength -Text $sampleValues -MaxLength 160)); leave suggested_edit empty."
}

function Test-AiSuggestedValueCanApplyToGroup {
    param([object[]]$Rows)

    if ($Rows.Count -le 1) { return $true }

    $observedValues = @(Get-UniqueExportRowValues -Rows $Rows -ColumnName "ObservedValue")
    $suggestedValues = @(Get-UniqueExportRowValues -Rows $Rows -ColumnName "SuggestedValue" -NonBlankOnly)
    if ($observedValues.Count -gt 1) { return $false }
    if ($suggestedValues.Count -gt 1) { return $false }
    return $true
}

function Get-AiExportGroupKey {
    param([object]$Row)

    $parts = @(
        [string](Get-RowValue -Row $Row -Name "TableName"),
        [string](Get-RowValue -Row $Row -Name "FieldName"),
        [string](Get-RowValue -Row $Row -Name "RuleName"),
        [string](Get-RowValue -Row $Row -Name "Message"),
        [string](Get-RowValue -Row $Row -Name "SuggestedEdit"),
        [string](Get-RowValue -Row $Row -Name "RuleSuggestedEdit"),
        [string](Get-RowValue -Row $Row -Name "_RuleContext")
    )
    return [string]::Join("|", $parts)
}

function Add-AiSuggestedEditsToExportRows {
    param(
        [System.Data.DataTable]$ExportRows,
        [int]$ProgressStart,
        [int]$ProgressEnd
    )

    if ($ExportRows.Rows.Count -eq 0) { return }

    $groups = @($ExportRows.Rows | Group-Object { Get-AiExportGroupKey -Row $_ })
    $progressSpan = [Math]::Max(1, $ProgressEnd - $ProgressStart)
    Set-AppProgress -Status "AI guidance: $($groups.Count) grouped request(s) for $($ExportRows.Rows.Count) rows..." -Percent $ProgressStart

    $script:LastAiSuggestedEditGroupsApplied = 0
    $script:LastAiSuggestedEditGroupsSkipped = 0
    $script:LastAiSuggestedEditRowsApplied = 0
    $script:LastAiSuggestedEditRowsSkipped = 0
    $script:LastAiSuggestedEditManualGroupsUsed = 0
    $script:LastAiSuggestedEditManualRowsUsed = 0
    $script:LastAiSuggestedEditNoManualRetryGroups = 0
    $script:LastAiSuggestedEditNoManualRetryRows = 0
    $script:LastAiSuggestedEditNoManualSnippetGroups = 0
    $script:LastAiSuggestedEditNoManualSnippetRows = 0

    $groupIndex = 0
    foreach ($group in $groups) {
        Assert-GuideRunNotCanceled
        $groupIndex++
        $sample = $group.Group[0]
        $fieldName = [string](Get-RowValue -Row $sample -Name "FieldName")
        $ruleName = [string](Get-RowValue -Row $sample -Name "RuleName")
        $groupPercent = $ProgressStart + [int](($groupIndex / [double]$groups.Count) * $progressSpan)
        $statusText = "AI guidance group $groupIndex of $($groups.Count)"
        if (-not [string]::IsNullOrWhiteSpace($fieldName)) {
            $statusText += ": $fieldName"
        }
        if (-not [string]::IsNullOrWhiteSpace($ruleName)) {
            $statusText += " / $ruleName"
        }
        Set-AppProgress -Status $statusText -Percent $groupPercent

        $allowSuggestedValue = $false
        $aiSuggestion = $null
        try {
            $aiSuggestion = Get-AiProjectGuidanceForExportFinding `
                -TableName ([string](Get-RowValue -Row $sample -Name "TableName")) `
                -FieldName $fieldName `
                -RuleName $ruleName `
                -Message ([string](Get-RowValue -Row $sample -Name "Message")) `
                -SuggestedEdit ([string](Get-RowValue -Row $sample -Name "SuggestedEdit")) `
                -RuleSuggestedEdit ([string](Get-RowValue -Row $sample -Name "RuleSuggestedEdit")) `
                -SuggestedValue ([string](Get-RowValue -Row $sample -Name "SuggestedValue")) `
                -SuggestedValueSummary (Get-AiSuggestedValueSummaryForGroup -Rows ([object[]]$group.Group)) `
                -ObservedValue ([string](Get-RowValue -Row $sample -Name "ObservedValue")) `
                -RecordLabel ([string](Get-RowValue -Row $sample -Name "RecordLabel")) `
                -RuleContext ([string](Get-RowValue -Row $sample -Name "_RuleContext")) `
                -GroupCount $group.Count `
                -AllowSuggestedValue:$allowSuggestedValue
        }
        catch {
            $errorText = Limit-TextLength -Text $_.Exception.Message -MaxLength 700
            $script:LastAiSuggestedEditGroupsSkipped++
            $script:LastAiSuggestedEditRowsSkipped += $group.Count
            foreach ($row in $group.Group) {
                $row["AIMessage"] = "AI skipped this finding group. Reason: $errorText"
            }
            Set-AppProgress -Status "AI skipped $fieldName / ${ruleName}: $errorText" -Percent $groupPercent
            continue
        }

        $usedAiForGroup = $false
        $aiMessage = [string](Get-RowValue -Row $aiSuggestion -Name "AIMessage")
        if ([string]::IsNullOrWhiteSpace($aiMessage)) {
            $aiMessage = [string](Get-RowValue -Row $aiSuggestion -Name "Guidance")
        }
        $aiSuggestedEditValue = ""
        if ($allowSuggestedValue) {
            $aiSuggestedEditValue = Normalize-AiSuggestedEditValue (Get-RowValue -Row $aiSuggestion -Name "SuggestedEdit")
            if ([string]::IsNullOrWhiteSpace($aiSuggestedEditValue)) {
                $aiSuggestedEditValue = Normalize-AiSuggestedEditValue (Get-RowValue -Row $aiSuggestion -Name "SuggestedValue")
            }
        }
        $manualSnippetAvailable = if ($aiSuggestion.PSObject.Properties["ManualSnippetAvailable"]) { [bool]$aiSuggestion.ManualSnippetAvailable } else { $false }
        $manualSnippetIncluded = if ($aiSuggestion.PSObject.Properties["ManualSnippetIncluded"]) { [bool]$aiSuggestion.ManualSnippetIncluded } else { $false }
        $promptAttempt = if ($aiSuggestion.PSObject.Properties["PromptAttempt"]) { [string]$aiSuggestion.PromptAttempt } else { "" }

        foreach ($row in $group.Group) {
            if (-not [string]::IsNullOrWhiteSpace($aiMessage)) {
                $row["AIMessage"] = $aiMessage
                $usedAiForGroup = $true
            }
            if ($allowSuggestedValue -and -not [string]::IsNullOrWhiteSpace($aiSuggestedEditValue)) {
                $row["SuggestedEdit"] = $aiSuggestedEditValue
                $row["SuggestedValue"] = $aiSuggestedEditValue
                $usedAiForGroup = $true
            }
        }

        if ($usedAiForGroup) {
            $script:LastAiSuggestedEditGroupsApplied++
            $script:LastAiSuggestedEditRowsApplied += $group.Count
            if ($manualSnippetIncluded) {
                $script:LastAiSuggestedEditManualGroupsUsed++
                $script:LastAiSuggestedEditManualRowsUsed += $group.Count
            }
            elseif (-not $manualSnippetAvailable) {
                $script:LastAiSuggestedEditNoManualSnippetGroups++
                $script:LastAiSuggestedEditNoManualSnippetRows += $group.Count
            }
            if ($promptAttempt -eq "tiny-no-manual") {
                $script:LastAiSuggestedEditNoManualRetryGroups++
                $script:LastAiSuggestedEditNoManualRetryRows += $group.Count
            }
        }
        else {
            foreach ($row in $group.Group) {
                $row["AIMessage"] = "AI did not return guidance for this finding group."
            }
            $script:LastAiSuggestedEditGroupsSkipped++
            $script:LastAiSuggestedEditRowsSkipped += $group.Count
        }
    }
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Text
    )

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $encoding)
}

function Add-DatabaseSquareReportRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Check,
        [object]$PeriodNumber,
        [string]$ExpectedTable,
        [object]$ExpectedCount,
        [string]$ObservedTable,
        [object]$ObservedCount,
        [string]$Status,
        [string]$Notes
    )

    [void]$Rows.Add([pscustomobject]@{
        Check = $Check
        PeriodNumber = if ($null -eq $PeriodNumber) { "" } else { [string]$PeriodNumber }
        ExpectedTable = $ExpectedTable
        ExpectedCount = if ($null -eq $ExpectedCount) { "" } else { [string]$ExpectedCount }
        ObservedTable = $ObservedTable
        ObservedCount = if ($null -eq $ObservedCount) { "" } else { [string]$ObservedCount }
        Status = $Status
        Notes = $Notes
    })
}

function Add-DatabaseSquareCountRows {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [System.Collections.Generic.List[object]]$Rows,
        [string[]]$Tables,
        [int[]]$Periods,
        [string]$Check,
        [string]$ExpectedTable,
        [string]$ObservedTable,
        [scriptblock]$ExpectedCountScript,
        [scriptblock]$ObservedCountScript,
        [string]$Notes
    )

    if (-not (Test-TableAvailable -Tables $Tables -TableName $ExpectedTable) -or
        -not (Test-TableAvailable -Tables $Tables -TableName $ObservedTable)) {
        Add-DatabaseSquareReportRow `
            -Rows $Rows `
            -Check $Check `
            -PeriodNumber "" `
            -ExpectedTable $ExpectedTable `
            -ExpectedCount $null `
            -ObservedTable $ObservedTable `
            -ObservedCount $null `
            -Status "Not checked" `
            -Notes "Required table missing, so this square-count check could not run."
        return
    }

    foreach ($period in $Periods) {
        try {
            $expected = [int](& $ExpectedCountScript $period)
            $observed = [int](& $ObservedCountScript $period)
            $status = if ($expected -eq $observed) { "OK" } else { "Review" }
            Add-DatabaseSquareReportRow `
                -Rows $Rows `
                -Check $Check `
                -PeriodNumber $period `
                -ExpectedTable $ExpectedTable `
                -ExpectedCount $expected `
                -ObservedTable $ObservedTable `
                -ObservedCount $observed `
                -Status $status `
                -Notes $Notes
        }
        catch {
            Add-DatabaseSquareReportRow `
                -Rows $Rows `
                -Check $Check `
                -PeriodNumber $period `
                -ExpectedTable $ExpectedTable `
                -ExpectedCount $null `
                -ObservedTable $ObservedTable `
                -ObservedCount $null `
                -Status "Not checked" `
                -Notes ("Check failed: " + $_.Exception.Message)
        }
    }
}

function Get-DatabaseSquareReportRows {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $rows = New-Object System.Collections.Generic.List[object]
    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    try {
        $periods = [int[]]@(Get-PeriodNumbers -Connection $Connection)
    }
    catch {
        Add-DatabaseSquareReportRow `
            -Rows $rows `
            -Check "Measurement periods" `
            -PeriodNumber "" `
            -ExpectedTable "ProjectMeasurementPeriods" `
            -ExpectedCount $null `
            -ObservedTable "" `
            -ObservedCount $null `
            -Status "Not checked" `
            -Notes ("Measurement periods could not be read: " + $_.Exception.Message)
        return @($rows.ToArray())
    }

    if ($periods.Count -eq 0) {
        Add-DatabaseSquareReportRow `
            -Rows $rows `
            -Check "Measurement periods" `
            -PeriodNumber "" `
            -ExpectedTable "ProjectMeasurementPeriods" `
            -ExpectedCount $null `
            -ObservedTable "" `
            -ObservedCount $null `
            -Status "Not checked" `
            -Notes "No measurement periods were found, so square-count checks could not run."
        return @($rows.ToArray())
    }

    Add-DatabaseSquareCountRows `
        -Connection $Connection `
        -Rows $rows `
        -Tables $tables `
        -Periods $periods `
        -Check "PlotMeasurements match Plots" `
        -ExpectedTable "Plots" `
        -ObservedTable "PlotMeasurements" `
        -ExpectedCountScript { param($period) Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [Plots]$(Get-ActiveWhereClause -Connection $Connection -TableName 'Plots')" } `
        -ObservedCountScript { param($period) Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [PlotMeasurements] WHERE [PeriodNumber] = $period" } `
        -Notes "Each period should usually have one PlotMeasurements row for each active plot."

    Add-DatabaseSquareCountRows `
        -Connection $Connection `
        -Rows $rows `
        -Tables $tables `
        -Periods $periods `
        -Check "PlotCustomMeasurements match PlotMeasurements" `
        -ExpectedTable "PlotMeasurements" `
        -ObservedTable "PlotCustomMeasurements" `
        -ExpectedCountScript { param($period) Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [PlotMeasurements] WHERE [PeriodNumber] = $period" } `
        -ObservedCountScript { param($period) Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [PlotCustomMeasurements] AS c INNER JOIN [PlotMeasurements] AS m ON c.[PlotMeasKey] = m.[PlotMeasKey] WHERE m.[PeriodNumber] = $period" } `
        -Notes "Custom plot measurements should line up with plot measurements by period."

    Add-DatabaseSquareCountRows `
        -Connection $Connection `
        -Rows $rows `
        -Tables $tables `
        -Periods $periods `
        -Check "TreeMeasurements match Trees" `
        -ExpectedTable "Trees" `
        -ObservedTable "TreeMeasurements" `
        -ExpectedCountScript { param($period) Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [Trees]$(Get-ActiveWhereClause -Connection $Connection -TableName 'Trees')" } `
        -ObservedCountScript { param($period) Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [TreeMeasurements] WHERE [PeriodNumber] = $period" } `
        -Notes "TreeMeasurements should usually match the tree record count by period; dropped or unmeasured plots may explain real differences."

    Add-DatabaseSquareCountRows `
        -Connection $Connection `
        -Rows $rows `
        -Tables $tables `
        -Periods $periods `
        -Check "TreeCustomMeasurements match TreeMeasurements" `
        -ExpectedTable "TreeMeasurements" `
        -ObservedTable "TreeCustomMeasurements" `
        -ExpectedCountScript { param($period) Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [TreeMeasurements] WHERE [PeriodNumber] = $period" } `
        -ObservedCountScript { param($period) Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [TreeCustomMeasurements] AS c INNER JOIN [TreeMeasurements] AS m ON c.[TreeMeasKey] = m.[TreeMeasKey] WHERE m.[PeriodNumber] = $period" } `
        -Notes "Custom tree measurements should line up with tree measurements by period."

    $regenSquareCheck = "RegenCustomMeasurements match RegenMeasurements"
    if (-not (Test-TableAvailable -Tables $tables -TableName "RegenMeasurements") -or
        -not (Test-TableAvailable -Tables $tables -TableName "RegenCustomMeasurements")) {
        Add-DatabaseSquareReportRow `
            -Rows $rows `
            -Check $regenSquareCheck `
            -PeriodNumber "" `
            -ExpectedTable "RegenMeasurements" `
            -ExpectedCount $null `
            -ObservedTable "RegenCustomMeasurements" `
            -ObservedCount $null `
            -Status "Not checked" `
            -Notes "Required regen table missing, so this square-count check could not run."
    }
    else {
        try {
            $regenColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenMeasurements")
            $regenCustomColumns = @(Get-TableColumns -Connection $Connection -TableName "RegenCustomMeasurements")
            if (-not (Test-ColumnExists -Columns $regenColumns -Name "RegenMeasKey") -or
                -not (Test-ColumnExists -Columns $regenCustomColumns -Name "RegenMeasKey")) {
                Add-DatabaseSquareReportRow `
                    -Rows $rows `
                    -Check $regenSquareCheck `
                    -PeriodNumber "" `
                    -ExpectedTable "RegenMeasurements" `
                    -ExpectedCount $null `
                    -ObservedTable "RegenCustomMeasurements" `
                    -ObservedCount $null `
                    -Status "Not checked" `
                    -Notes "Both RegenMeasurements and RegenCustomMeasurements need RegenMeasKey for this square-count check."
            }
            else {
                foreach ($period in $periods) {
                    $regenCondition = Get-RegenMeasKeyPeriodCondition -PeriodNumber ([int]$period)
                    $regenCustomCondition = Get-RegenMeasKeyPeriodCondition -PeriodNumber ([int]$period)
                    $expected = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [RegenMeasurements] WHERE $regenCondition"
                    $observed = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [RegenCustomMeasurements] WHERE $regenCustomCondition"
                    Add-DatabaseSquareReportRow `
                        -Rows $rows `
                        -Check "RegenMeasurements count by RegenMeasKey" `
                        -PeriodNumber $period `
                        -ExpectedTable "RegenMeasurements" `
                        -ExpectedCount $expected `
                        -ObservedTable "Decoded RegenMeasKey period" `
                        -ObservedCount $expected `
                        -Status "OK" `
                        -Notes "Regen has no separate parent table like Trees. This row shows the RegenMeasurements table count for the period decoded from RegenMeasKey."
                    $status = if ($expected -eq $observed) { "OK" } else { "Review" }
                    Add-DatabaseSquareReportRow `
                        -Rows $rows `
                        -Check $regenSquareCheck `
                        -PeriodNumber $period `
                        -ExpectedTable "RegenMeasurements" `
                        -ExpectedCount $expected `
                        -ObservedTable "RegenCustomMeasurements" `
                        -ObservedCount $observed `
                        -Status $status `
                        -Notes "Custom regen measurements should line up with regen measurements by period decoded from RegenMeasKey. Example: P005337-01-3-122-040 = plot 5337, period 1, minor plot 3, species 122, IDBH 40."
                }
            }
        }
        catch {
            Add-DatabaseSquareReportRow `
                -Rows $rows `
                -Check $regenSquareCheck `
                -PeriodNumber "" `
                -ExpectedTable "RegenMeasurements" `
                -ExpectedCount $null `
                -ObservedTable "RegenCustomMeasurements" `
                -ObservedCount $null `
                -Status "Not checked" `
                -Notes ("Regen key square-count check failed: " + $_.Exception.Message)
        }
    }

    return @($rows.ToArray())
}

function Get-MissingTreeMeasurementReportRows {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $rows = New-Object System.Collections.Generic.List[object]
    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    foreach ($requiredTable in @("Trees", "TreeMeasurements", "ProjectMeasurementPeriods")) {
        if (-not (Test-TableAvailable -Tables $tables -TableName $requiredTable)) {
            [void]$rows.Add([pscustomobject]@{
                Status = "Not checked"
                PlotNumber = ""
                TreeNumber = ""
                SpeciesCode = ""
                MinorPlot = ""
                TreeID = ""
                PeriodNumber = ""
                RecordLabel = ""
                Notes = "Required table '$requiredTable' was not found, so missing tree measurements could not be checked."
            })
            return @($rows.ToArray())
        }
    }

    $treeColumns = @(Get-TableColumns -Connection $Connection -TableName "Trees")
    $treeMeasColumns = @(Get-TableColumns -Connection $Connection -TableName "TreeMeasurements")
    $projectPeriodColumns = @(Get-TableColumns -Connection $Connection -TableName "ProjectMeasurementPeriods")
    foreach ($columnName in @("TreeID", "PlotNumber", "TreeNumber")) {
        if (-not (Test-ColumnExists -Columns $treeColumns -Name $columnName)) {
            [void]$rows.Add([pscustomobject]@{
                Status = "Not checked"
                PlotNumber = ""
                TreeNumber = ""
                SpeciesCode = ""
                MinorPlot = ""
                TreeID = ""
                PeriodNumber = ""
                RecordLabel = ""
                Notes = "Trees.$columnName was not found, so missing tree measurements could not be checked."
            })
            return @($rows.ToArray())
        }
    }
    foreach ($columnName in @("TreeID", "PeriodNumber")) {
        if (-not (Test-ColumnExists -Columns $treeMeasColumns -Name $columnName)) {
            [void]$rows.Add([pscustomobject]@{
                Status = "Not checked"
                PlotNumber = ""
                TreeNumber = ""
                SpeciesCode = ""
                MinorPlot = ""
                TreeID = ""
                PeriodNumber = ""
                RecordLabel = ""
                Notes = "TreeMeasurements.$columnName was not found, so missing tree measurements could not be checked."
            })
            return @($rows.ToArray())
        }
    }
    if (-not (Test-ColumnExists -Columns $projectPeriodColumns -Name "PeriodNumber")) {
        [void]$rows.Add([pscustomobject]@{
            Status = "Not checked"
            PlotNumber = ""
            TreeNumber = ""
            SpeciesCode = ""
            MinorPlot = ""
            TreeID = ""
            PeriodNumber = ""
            RecordLabel = ""
            Notes = "ProjectMeasurementPeriods.PeriodNumber was not found, so missing tree measurements could not be checked."
        })
        return @($rows.ToArray())
    }

    try {
        $periodScope = Get-SelectedPeriodAliasCondition -Alias "p"
        $periodScopeWhere = if ([string]::IsNullOrWhiteSpace($periodScope)) { "1 = 1" } else { $periodScope }
        $minorPlotColumn = Get-MinorPlotColumnName -Columns $treeColumns
        $minorPlotSelect = if ([string]::IsNullOrWhiteSpace($minorPlotColumn)) { "'' AS [MinorPlot]" } else { "t.$(Quote-Name $minorPlotColumn) AS [MinorPlot]" }
        $speciesCodeSelect = if (Test-ColumnExists -Columns $treeColumns -Name "SpeciesCode") { "(t.[SpeciesCode] & '') AS [SpeciesCode]" } else { "'' AS [SpeciesCode]" }
        $missingRows = Get-DataTable -Connection $Connection -Sql @"
SELECT
    t.[PlotNumber],
    t.[TreeNumber],
    $speciesCodeSelect,
    $minorPlotSelect,
    t.[TreeID],
    p.[PeriodNumber],
    (t.[PlotNumber] & '/' & t.[TreeNumber] & ' P' & p.[PeriodNumber]) AS [RecordLabel]
FROM [Trees] AS t, [ProjectMeasurementPeriods] AS p
WHERE $periodScopeWhere
  AND NOT EXISTS (
        SELECT 1
        FROM [TreeMeasurements] AS tm
        WHERE tm.[TreeID] = t.[TreeID]
          AND tm.[PeriodNumber] = p.[PeriodNumber]
    )
  AND EXISTS (
        SELECT 1
        FROM [TreeMeasurements] AS tm2
        WHERE tm2.[TreeID] = t.[TreeID]
    )
ORDER BY t.[PlotNumber], t.[TreeNumber], p.[PeriodNumber]
"@
    }
    catch {
        [void]$rows.Add([pscustomobject]@{
            Status = "Not checked"
            PlotNumber = ""
            TreeNumber = ""
            SpeciesCode = ""
            MinorPlot = ""
            TreeID = ""
            PeriodNumber = ""
            RecordLabel = ""
            Notes = ("Missing tree measurement report failed: " + $_.Exception.Message)
        })
        return @($rows.ToArray())
    }

    if ($missingRows.Rows.Count -eq 0) {
        [void]$rows.Add([pscustomobject]@{
            Status = "OK"
            PlotNumber = ""
            TreeNumber = ""
            SpeciesCode = ""
            MinorPlot = ""
            TreeID = ""
            PeriodNumber = ""
            RecordLabel = ""
            Notes = "No trees with at least one measurement are missing TreeMeasurements rows for another project measurement period."
        })
        return @($rows.ToArray())
    }

    foreach ($row in $missingRows.Rows) {
        [void]$rows.Add([pscustomobject]@{
            Status = "Missing"
            PlotNumber = Get-DataRowText -Row $row -ColumnName "PlotNumber"
            TreeNumber = Get-DataRowText -Row $row -ColumnName "TreeNumber"
            SpeciesCode = Get-DataRowText -Row $row -ColumnName "SpeciesCode"
            MinorPlot = Get-DataRowText -Row $row -ColumnName "MinorPlot"
            TreeID = Get-DataRowText -Row $row -ColumnName "TreeID"
            PeriodNumber = Get-DataRowText -Row $row -ColumnName "PeriodNumber"
            RecordLabel = Get-DataRowText -Row $row -ColumnName "RecordLabel"
            Notes = "Tree has at least one measurement, but is missing a TreeMeasurements row for this project measurement period."
        })
    }

    return @($rows.ToArray())
}

function Get-ReportStatusCount {
    param(
        [object[]]$Rows,
        [string[]]$Statuses
    )

    $count = 0
    foreach ($row in $Rows) {
        $status = [string](Get-RowValue -Row $row -Name "Status")
        if ($Statuses -contains $status) { $count++ }
    }
    return $count
}

function Get-CurrentGuideRunElapsedText {
    $startedAt = $script:GuideRunStartedAt
    $settings = $script:GuideRunSettings
    if (($null -eq $startedAt) -and ($null -ne $settings) -and $settings.PSObject.Properties["StartedAt"]) {
        $startedAt = ConvertTo-GuideRunDateTime $settings.StartedAt
    }

    if ($null -eq $startedAt) { return "" }
    return Format-ElapsedText -Elapsed ((Get-Date) - $startedAt)
}

function Set-RunDataRowValue {
    param(
        [object[]]$Rows,
        [string]$Section,
        [string]$Item,
        [object]$Value
    )

    foreach ($row in $Rows) {
        if ([string](Get-RowValue -Row $row -Name "Section") -ne $Section) { continue }
        if ([string](Get-RowValue -Row $row -Name "Item") -ne $Item) { continue }
        $row.Value = if ($null -eq $Value) { "" } else { [string]$Value }
        return
    }
}

function Get-RunDataRows {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$ExportPath,
        [int]$AuditRowCount,
        [int]$ExportedRowCount,
        [object[]]$DatabaseSquareRows,
        [object[]]$MissingTreeRows,
        [object]$RunCounts = $null,
        [switch]$UseAiProjectGuidance
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $addRow = {
        param([string]$Section, [string]$Item, [object]$Value)
        $text = if ($null -eq $Value) { "" } else { [string]$Value }
        [void]$rows.Add([pscustomobject]@{
            Section = $Section
            Item = $Item
            Value = $text
        })
    }

    $settings = $script:GuideRunSettings
    $sourcePath = if ($null -ne $settings) { [string]$settings.SourcePath } else { [string]$script:DbPath }
    $manualPath = if ($null -ne $settings) { [string]$settings.ProjectManualPath } else { [string]$script:ProjectManualPath }
    $manualLoaded = if ($null -ne $settings) { -not [string]::IsNullOrWhiteSpace([string]$settings.ProjectManualText) } else { -not [string]::IsNullOrWhiteSpace([string]$script:ProjectManualText) }
    $aiEndpoint = if ($null -ne $settings) { [string]$settings.AiEndpoint } else { [string]$script:AiDefaultEndpoint }
    $aiModel = if ($null -ne $settings) { [string]$settings.AiModel } else { [string]$script:AiDefaultModel }
    $aiHelperOn = if ($null -ne $settings) { [bool]$settings.AiHelperOn } else { Test-AiHelperIsOn }
    $compactAi = if ($null -ne $settings -and $settings.PSObject.Properties["AiCompactExport"]) { [bool]$settings.AiCompactExport } else { Test-AiCompactExportSelected }
    $runLogPath = if ($null -ne $settings -and $settings.PSObject.Properties["RunLogPath"]) { [string]$settings.RunLogPath } else { "" }

    $endpointHost = ""
    try {
        $endpointUri = [System.Uri]::new((Get-ValidatedAiEndpoint -Endpoint $aiEndpoint))
        $endpointHost = $endpointUri.Host
    }
    catch {
        $endpointHost = "Not available"
    }

    & $addRow "Run" "App" $script:AppName
    & $addRow "Run" "Build" $script:AppVersion
    & $addRow "Run" "Export created" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    & $addRow "Run" "App run time" (Get-CurrentGuideRunElapsedText)
    & $addRow "Run" "Process bitness" $(if ([Environment]::Is64BitProcess) { "64-bit" } else { "32-bit" })
    & $addRow "Run" "Source database" $sourcePath
    & $addRow "Run" "Source database file" $(if ([string]::IsNullOrWhiteSpace($sourcePath)) { "" } else { [System.IO.Path]::GetFileName($sourcePath) })
    & $addRow "Run" "Export workbook" $ExportPath
    & $addRow "Run" "Run log" $runLogPath
    & $addRow "Run" "Run modifies original database" "No - checks run on a temporary database copy"

    & $addRow "Manual" "Project manual" $manualPath
    & $addRow "Manual" "Project manual loaded" $(if ($manualLoaded) { "Yes" } else { "No" })

    & $addRow "Options" "Range checks enabled" $(if (Test-RangeChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "AppColumns active-field filter" "Only fields marked Active in AppColumns are reviewed for data-entry errors; excluded fields such as per-acre expansion, CalcSiteIndex, and GMP are skipped."
    & $addRow "Options" "Required plot rules" "Active elevation, aspect, slope, UTM, measurement date, crew, stand class/age, and stockability fields are required for PlotStatus 1 or 2; PlotRemarks is required for PlotStatus 7; PlotStatus progression is checked across all periods when PlotMeasurements.PlotStatus is active."
    & $addRow "Options" "Split DBH/height/TreeHistory history columns" $(if (Test-SplitHistoryColumnsSelected) { "Yes" } else { "No" })
    & $addRow "Options" "Max IDBH" (Get-ControlDecimalValue -Control $dbhMax -DefaultValue 500)
    & $addRow "Options" "Max height" (Get-ControlDecimalValue -Control $heightMax -DefaultValue 150)
    & $addRow "Options" "Max regen stems" (Get-StemCountMaximumValue)
    & $addRow "Options" "Check tree StemCount" $(if (Test-TreeStemCountChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Cleaning period scope" (Get-CleaningPeriodScopeDescription)
    & $addRow "Options" "Limit normal checks to measurement periods" $(if (Test-CleaningPeriodScopeEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Single measurement project" $(if (Test-SingleMeasurementProjectEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Current measurement period number" (Get-PeriodScopeCurrentPeriodValue)
    & $addRow "Options" "Previous measurement period number" (Get-PeriodScopePastPeriodValue)
    & $addRow "Options" "Require DBH for new mortality TreeHistory 2/3" $(if (Test-NewMortalityIdbhChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Require DBH for old mortality TreeHistory 7" $(if (Test-OldMortalityIdbhChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Require height for new mortality TreeHistory 2/3" $(if (Test-NewMortalityHeightChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Require height for old mortality TreeHistory 7" $(if (Test-OldMortalityHeightChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Flag entered timber height for Problem 127/legacy 74 (broken/missing top)" $(if (Test-Problem127HeightChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Flag entered timber height for Problem 128/legacy 75 (dead top)" $(if (Test-Problem128HeightChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Flag entered timber height for Problem 123/legacy 72 (lean > 15 degrees)" $(if (Test-Problem123HeightChecksEnabled) { "Yes" } else { "No" })
    & $addRow "Options" "Total Height Protocol" (Get-TotalHeightProtocolDisplayText)
    & $addRow "Options" "Total Height subsample minimum count" (Get-HeightSubsampleMinimumCount)
    & $addRow "Options" "Total Height subsample minimum eligible IDBH" (Get-HeightSubsampleMinimumIdbh)
    & $addRow "Options" "Total Height subsample all heights at/above IDBH" $(if (Test-HeightSubsampleAllAtOrAboveEnabled) { Get-HeightSubsampleAllAtOrAboveIdbh } else { "Off" })
    & $addRow "Options" "Rare species requiring 100% live-tree height" (Get-ControlTextValue -Control $heightRareSpeciesCodesBox -DefaultValue "" -SettingName "HeightRareSpeciesCodesText")
    & $addRow "Options" "Minor plots where tree height is required" (Get-ControlTextValue -Control $heightRequiredMinorPlotsBox -DefaultValue $script:DefaultHeightRequiredMinorPlots -SettingName "HeightRequiredMinorPlotsText")
    & $addRow "Options" "Regen timber seedling minor plots" (Get-ControlTextValue -Control $regenTimberSeedlingMinorPlotsBox -DefaultValue $script:DefaultRegenTimberSeedlingMinorPlots -SettingName "RegenTimberSeedlingMinorPlotsText")
    & $addRow "Options" "Regen timber sapling IDBH 20 minor plots" (Get-ControlTextValue -Control $regenTimberSapling20MinorPlotsBox -DefaultValue $script:DefaultRegenTimberSapling20MinorPlots -SettingName "RegenTimberSapling20MinorPlotsText")
    & $addRow "Options" "Regen timber sapling IDBH 40 minor plots" (Get-ControlTextValue -Control $regenTimberSapling40MinorPlotsBox -DefaultValue $script:DefaultRegenTimberSapling40MinorPlots -SettingName "RegenTimberSapling40MinorPlotsText")
    & $addRow "Options" "Regen woodland seedling minor plots" (Get-ControlTextValue -Control $regenWoodlandSeedlingMinorPlotsBox -DefaultValue $script:DefaultRegenWoodlandSeedlingMinorPlots -SettingName "RegenWoodlandSeedlingMinorPlotsText")
    & $addRow "Options" "Regen woodland sapling IDBH 20 minor plots" (Get-ControlTextValue -Control $regenWoodlandSapling20MinorPlotsBox -DefaultValue $script:DefaultRegenWoodlandSapling20MinorPlots -SettingName "RegenWoodlandSapling20MinorPlotsText")
    & $addRow "Options" "Regen woodland sapling IDBH 40 minor plots" (Get-ControlTextValue -Control $regenWoodlandSapling40MinorPlotsBox -DefaultValue $script:DefaultRegenWoodlandSapling40MinorPlots -SettingName "RegenWoodlandSapling40MinorPlotsText")
    & $addRow "Options" "Max period-to-period IDBH jump" (Get-ControlDecimalValue -Control $dbhGrowthMax -DefaultValue 100)
    & $addRow "Options" "Max period-to-period height jump" (Get-ControlDecimalValue -Control $heightGrowthMax -DefaultValue 20)
    & $addRow "Options" "Woodland species excluded from shrinking IDBH checks" (Get-ControlTextValue -Control $woodlandSpeciesCodesBox -DefaultValue $script:DefaultWoodlandSpeciesCodes -SettingName "WoodlandSpeciesCodesText")

    & $addRow "AI" "AI helper on" $(if ($aiHelperOn) { "Yes" } else { "No" })
    & $addRow "AI" "AI guidance used for export" $(if ($UseAiProjectGuidance) { "Yes" } else { "No" })
    & $addRow "AI" "Compact AI export" $(if ($compactAi) { "Yes" } else { "No" })
    & $addRow "AI" "Approved endpoint host" $endpointHost
    & $addRow "AI" "Model/deployment" $aiModel
    & $addRow "AI" "API key exported" "No"

    & $addRow "Findings" "InventoryCleanAudit rows read" $AuditRowCount
    & $addRow "Findings" "Excel rows exported" $ExportedRowCount
    & $addRow "Findings" "Reviewed-clean field rows added" $script:LastReviewedCleanExportRows
    & $addRow "Findings" "Database Square review rows" (Get-ReportStatusCount -Rows $DatabaseSquareRows -Statuses @("Review"))
    & $addRow "Findings" "Missing Trees rows" (Get-ReportStatusCount -Rows $MissingTreeRows -Statuses @("Missing"))

    if ($null -ne $RunCounts) {
        foreach ($name in @("WorkbookCounts", "RScriptReference", "InvalidCodes", "Ranges", "PlotStatusProgression", "InactivePlotData", "InactivePlotTreeData", "InactivePlotRegenData", "Remeasurement", "TreeHistoryConversion", "Ingrowth")) {
            if ($RunCounts.PSObject.Properties[$name]) {
                & $addRow "Audit counts" $name $RunCounts.$name
            }
        }
    }

    $topPerformanceEntries = @(Get-GuideRunTopPerformanceEntries -MaxCount 8)
    if ($topPerformanceEntries.Count -gt 0) {
        & $addRow "Performance" "Timing detail" "Detailed step timings and slow SQL previews are saved in the text run log."
        $performanceIndex = 0
        foreach ($entry in $topPerformanceEntries) {
            $performanceIndex++
            $text = "$($entry.Duration)"
            if (-not [string]::IsNullOrWhiteSpace([string]$entry.Count)) { $text += "; $($entry.Count)" }
            if ([string]$entry.Status -ne "completed") { $text += "; status=$($entry.Status)" }
            & $addRow "Performance" ("Slow step $performanceIndex - $($entry.Name)") $text
        }
        $slowSqlCount = if ($null -ne $script:GuideRunSlowSqlEntries) { $script:GuideRunSlowSqlEntries.Count } else { 0 }
        & $addRow "Performance" "Slow SQL/failure entries in run log" $slowSqlCount
    }

    if ($UseAiProjectGuidance) {
        & $addRow "AI results" "AI guidance groups applied" $script:LastAiSuggestedEditGroupsApplied
        & $addRow "AI results" "AI guidance groups fallback/skipped" $script:LastAiSuggestedEditGroupsSkipped
        & $addRow "AI results" "AI guidance rows applied" $script:LastAiSuggestedEditRowsApplied
        & $addRow "AI results" "AI guidance fallback rows" $script:LastAiSuggestedEditRowsSkipped
        & $addRow "AI results" "Uploaded manual groups used" $script:LastAiSuggestedEditManualGroupsUsed
        & $addRow "AI results" "Uploaded manual rows used" $script:LastAiSuggestedEditManualRowsUsed
        & $addRow "AI results" "No-manual retry groups" $script:LastAiSuggestedEditNoManualRetryGroups
        & $addRow "AI results" "No matching manual snippet groups" $script:LastAiSuggestedEditNoManualSnippetGroups
    }

    try {
        $targetTables = @(Get-CfiWorkbookTargetTables -Connection $Connection)
        foreach ($tableName in $targetTables) {
            try {
                $count = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM $(Quote-Name $tableName)"
                & $addRow "Source table rows" $tableName $count
            }
            catch {
                & $addRow "Source table rows" $tableName "Could not count rows"
            }
        }
    }
    catch {
    }

    return @($rows.ToArray())
}

function Export-AuditWorkbook {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$Path,
        [switch]$UseAiProjectGuidance,
        [object]$RunCounts = $null,
        [int]$ProgressStart = 15,
        [int]$ProgressEnd = 95
    )

    $script:PeriodScopeConditionCache = @{}
    $script:AppColumnReviewMetadataCache = @{}
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Reading audit findings..." -Percent $ProgressStart
    $exportReadStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
    if (-not ($tables -contains "InventoryCleanAudit")) {
        throw "InventoryCleanAudit was not found. Run first."
    }

    $auditColumns = @(Get-TableColumns -Connection $Connection -TableName "InventoryCleanAudit")
    $optionalColumns = @("SourceRowId", "PlotNumber", "TreeNumber", "SpeciesCode", "MinorPlot", "PeriodNumber", "TreeKey", "RegenMeasKey", "PlotStatus", "PlotRemarks", "RegenRemarks")
    $selectColumns = New-Object System.Collections.Generic.List[string]
    foreach ($columnName in @("AuditId", "TableName", "RuleName", "RecordLabel", "FieldName")) {
        [void]$selectColumns.Add((Quote-Name $columnName))
    }
    foreach ($columnName in $optionalColumns) {
        if (Test-ColumnExists -Columns $auditColumns -Name $columnName) {
            [void]$selectColumns.Add((Quote-Name $columnName))
        }
    }
    foreach ($columnName in @("ObservedValue", "Message")) {
        [void]$selectColumns.Add((Quote-Name $columnName))
    }
    $selectList = [string]::Join(", ", [string[]]$selectColumns.ToArray())
    $auditRows = Get-DataTable -Connection $Connection -Sql "SELECT $selectList FROM [InventoryCleanAudit] ORDER BY [FieldName], [TableName], [RuleName], [AuditId]"
    $exportReadStopwatch.Stop()
    Add-GuideRunPerformanceEntry -Name "Export: read audit findings" -Elapsed $exportReadStopwatch.Elapsed -CountText "$($auditRows.Rows.Count) audit row(s)"

    $progressSpan = [Math]::Max(1, $ProgressEnd - $ProgressStart)
    $rowProgressEnd = $ProgressStart + [int]($progressSpan * 0.70)
    $aiProgressEnd = $rowProgressEnd
    $script:LastAiSuggestedEditGroupsApplied = 0
    $script:LastAiSuggestedEditGroupsSkipped = 0
    $script:LastAiSuggestedEditRowsApplied = 0
    $script:LastAiSuggestedEditRowsSkipped = 0
    $script:LastAiSuggestedEditManualGroupsUsed = 0
    $script:LastAiSuggestedEditManualRowsUsed = 0
    $script:LastAiSuggestedEditNoManualRetryGroups = 0
    $script:LastAiSuggestedEditNoManualRetryRows = 0
    $script:LastAiSuggestedEditNoManualSnippetGroups = 0
    $script:LastAiSuggestedEditNoManualSnippetRows = 0
    if ($UseAiProjectGuidance) {
        $rowProgressEnd = $ProgressStart + [int]($progressSpan * 0.42)
        $aiProgressEnd = $ProgressStart + [int]($progressSpan * 0.72)
    }
    Set-AppProgress -Status "Preparing $($auditRows.Rows.Count) export rows..." -Percent $ProgressStart

    $splitHistoryColumns = Test-SplitHistoryColumnsSelected
    $headers = @("AuditId", "TableName", "RuleName", "ReportCurrentPeriodNumber", "ReportCurrentPeriodErrorCount", "ReportPreviousPeriodNumber", "ReportPreviousPeriodErrorCount", "RecordLabel", "FieldName", "PlotNumber", "TreeNumber", "SpeciesCode", "MinorPlot", "PeriodNumber", "PeriodScope", "TreeKey", "RegenMeasKey", "PlotStatus", "PlotRemarks", "RegenRemarks", "TreeHistory", "TreeRemarks", "SourceRowId", "ObservedValue", "RecordedValue", "Message", "VerificationSql", "SuggestedEdit", "AIMessage", "SuggestedValue", "RuleSuggestedEdit", "_RuleContext")
    if ($splitHistoryColumns) {
        $headers += Get-SplitHistoryBaseExportHeaders
    }
    $exportRows = New-Object System.Data.DataTable
    foreach ($header in $headers) { [void]$exportRows.Columns.Add($header) }

    if ($UseAiProjectGuidance -and $auditRows.Rows.Count -gt 0) {
        Assert-AiProjectGuidanceReady
    }

    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Indexing identity columns for export..." -Percent $ProgressStart
    $identityLookupMap = Get-AuditIdentityLookupMap -Connection $Connection -AuditRows $auditRows
    $verificationSqlCache = @{}
    $verificationSqlTestCache = @{}
    $ruleContextCache = @{}
    $treeHistoryTimelineCache = @{}
    $detailCurrentPeriodNumber = Get-PeriodScopeCurrentPeriodValue
    $detailPreviousPeriodNumber = Get-PeriodScopePastPeriodValue

    Assert-GuideRunNotCanceled
    $exportRowStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $rowIndex = 0
    foreach ($row in $auditRows.Rows) {
        Assert-GuideRunNotCanceled
        $rowIndex++
        $rowPercent = $ProgressStart
        if ($auditRows.Rows.Count -gt 0) {
            $rowPercent = $ProgressStart + [int](($rowIndex / [double]$auditRows.Rows.Count) * [Math]::Max(1, ($rowProgressEnd - $ProgressStart)))
        }
        if (($rowIndex -eq 1) -or ($rowIndex -eq $auditRows.Rows.Count) -or (($rowIndex % 25) -eq 0)) {
            Set-AppProgress -Status "Preparing export row $rowIndex of $($auditRows.Rows.Count)..." -Percent $rowPercent
        }

        $newRow = $exportRows.NewRow()
        foreach ($header in @("AuditId", "TableName", "RuleName", "RecordLabel", "FieldName", "ObservedValue")) {
            $newRow[$header] = Get-DataRowText -Row $row -ColumnName $header
        }
        $identity = Get-AuditIdentityForRowFast -Connection $Connection -Row $row -IdentityLookupMap $identityLookupMap
        $recordedValue = Get-DataRowText -Row $row -ColumnName "ObservedValue"
        $treeHistoryTimeline = Get-TreeHistoryTimelineRecordedValue `
            -Connection $Connection `
            -Row $row `
            -Identity $identity `
            -TimelineCache $treeHistoryTimelineCache
        if (-not [string]::IsNullOrWhiteSpace($treeHistoryTimeline)) {
            $recordedValue = $treeHistoryTimeline
        }
        $newRow["RecordedValue"] = $recordedValue
        if ($splitHistoryColumns) {
            Set-SplitHistoryExportColumns `
                -Table $exportRows `
                -Row $newRow `
                -ObservedValue (Get-DataRowText -Row $row -ColumnName "ObservedValue") `
                -RecordedValue $recordedValue
        }
        $newRow["Message"] = Get-DataRowText -Row $row -ColumnName "Message"
        foreach ($header in @("PlotNumber", "TreeNumber", "SpeciesCode", "MinorPlot", "PeriodNumber", "TreeKey", "RegenMeasKey", "PlotStatus", "PlotRemarks", "RegenRemarks", "TreeHistory", "TreeRemarks", "SourceRowId")) {
            $newRow[$header] = [string]$identity[$header]
        }
        $newRow["PeriodScope"] = Get-ReportPeriodScopeLabel -Row $newRow -CurrentPeriodNumber $detailCurrentPeriodNumber -PreviousPeriodNumber $detailPreviousPeriodNumber
        if ([string]::IsNullOrWhiteSpace([string]$newRow["PlotStatus"])) {
            $plotStatusFromObserved = Get-ObservedKeyValue -ObservedValue (Get-DataRowText -Row $row -ColumnName "ObservedValue") -KeyName "PlotStatus"
            if (-not [string]::IsNullOrWhiteSpace($plotStatusFromObserved)) {
                $newRow["PlotStatus"] = $plotStatusFromObserved
            }
        }
        $verificationKey = [string]::Join("|", @(
            (Get-DataRowText -Row $row -ColumnName "TableName"),
            (Get-DataRowText -Row $row -ColumnName "RuleName"),
            (Get-DataRowText -Row $row -ColumnName "FieldName")
        ))
        if (-not $verificationSqlCache.ContainsKey($verificationKey)) {
            $rawVerificationSql = Get-VerificationSqlForAuditRow -Connection $Connection -Row $row
            $verificationSqlCache[$verificationKey] = Remove-VerificationSqlTemporaryColumns -Sql $rawVerificationSql
        }
        if (-not $verificationSqlTestCache.ContainsKey($verificationKey)) {
            $verificationFieldName = Get-DataRowText -Row $row -ColumnName "FieldName"
            $verificationRuleName = Get-DataRowText -Row $row -ColumnName "RuleName"
            $verificationLabel = "$verificationFieldName / $verificationRuleName"
            Assert-VerificationSqlHasNoTemporaryColumns -Sql ([string]$verificationSqlCache[$verificationKey]) -Label $verificationLabel
            Assert-VerificationSqlRuns -Connection $Connection -Sql ([string]$verificationSqlCache[$verificationKey]) -Label $verificationLabel
            $verificationSqlTestCache[$verificationKey] = $true
        }
        $newRow["VerificationSql"] = [string]$verificationSqlCache[$verificationKey]
        $ruleSuggestedEdit = Get-SuggestedEditForAuditRow -Row $row
        $suggestedValue = Get-SuggestedValueForAuditRow -Row $row
        $newRow["SuggestedEdit"] = $suggestedValue
        $newRow["AIMessage"] = ""
        $newRow["SuggestedValue"] = $suggestedValue
        $newRow["RuleSuggestedEdit"] = $ruleSuggestedEdit
        if ($UseAiProjectGuidance) {
            if (-not $ruleContextCache.ContainsKey($verificationKey)) {
                $ruleContextCache[$verificationKey] = Get-TemplateRuleContextForAuditRow -Connection $Connection -Row $row
            }
            $newRow["_RuleContext"] = [string]$ruleContextCache[$verificationKey]
        }
        else {
            $newRow["_RuleContext"] = ""
        }
        [void]$exportRows.Rows.Add($newRow)
    }
    $exportRowStopwatch.Stop()
    Add-GuideRunPerformanceEntry -Name "Export: prepare detailed rows" -Elapsed $exportRowStopwatch.Elapsed -CountText "$($exportRows.Rows.Count) export row(s)"

    if ($UseAiProjectGuidance) {
        Assert-GuideRunNotCanceled
        [void](Invoke-GuideTimedStep -Name "Export: AI guidance groups" -ScriptBlock {
            Add-AiSuggestedEditsToExportRows `
                -ExportRows $exportRows `
                -ProgressStart $rowProgressEnd `
                -ProgressEnd $aiProgressEnd
        })
    }

    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Adding reviewed-clean field tabs..." -Percent ($ProgressStart + [int]($progressSpan * 0.72))
    [void](Invoke-GuideTimedStep -Name "Export: add reviewed-clean rows" -ScriptBlock {
        [void](Add-ReviewedCleanExportRows -Connection $Connection -ExportRows $exportRows)
        Set-PeriodScopeLabelsForRows -Rows @($exportRows.Rows) -CurrentPeriodNumber $detailCurrentPeriodNumber -PreviousPeriodNumber $detailPreviousPeriodNumber
    })

    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Building database square and missing tree reports..." -Percent ($ProgressStart + [int]($progressSpan * 0.73))
    $projectReportStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $databaseSquareRows = @(Get-DatabaseSquareReportRows -Connection $Connection)
    $missingTreeRows = @(Get-MissingTreeMeasurementReportRows -Connection $Connection)
    $projectReportStopwatch.Stop()
    Add-GuideRunPerformanceEntry -Name "Export: database square and missing tree reports" -Elapsed $projectReportStopwatch.Elapsed -CountText "$($databaseSquareRows.Count + $missingTreeRows.Count) report row(s)"
    $runDataRows = @(Get-RunDataRows `
        -Connection $Connection `
        -ExportPath $Path `
        -AuditRowCount $auditRows.Rows.Count `
        -ExportedRowCount $exportRows.Rows.Count `
        -DatabaseSquareRows $databaseSquareRows `
        -MissingTreeRows $missingTreeRows `
        -RunCounts $RunCounts `
        -UseAiProjectGuidance:$UseAiProjectGuidance)

    Set-AppProgress -Status "Building workbook sheets..." -Percent ($ProgressStart + [int]($progressSpan * 0.75))
    $sheetBuildStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $usedSheetNames = @{}
    $worksheets = New-Object System.Collections.Generic.List[object]
    $projectCheckWorksheets = New-Object System.Collections.Generic.List[object]
    $reviewWorksheets = New-Object System.Collections.Generic.List[object]
    $passedWorksheets = New-Object System.Collections.Generic.List[object]
    $summaryRows = New-Object System.Collections.Generic.List[object]
    $summarySheetName = Get-SafeWorksheetName -Name "Summary" -UsedNames $usedSheetNames
    $runDataSheetName = Get-SafeWorksheetName -Name "Run Data" -UsedNames $usedSheetNames
    $databaseSquareSheetName = Get-SafeWorksheetName -Name "Database Square" -UsedNames $usedSheetNames
    $missingTreesSheetName = Get-SafeWorksheetName -Name "Missing Trees" -UsedNames $usedSheetNames
    $summaryCurrentPeriodNumber = Get-PeriodScopeCurrentPeriodValue
    $summaryPreviousPeriodNumber = Get-PeriodScopePastPeriodValue
    $summaryCurrentPeriodDisplay = if ($summaryCurrentPeriodNumber -gt 0) { [string]$summaryCurrentPeriodNumber } else { "" }
    $summaryPreviousPeriodDisplay = if ($summaryPreviousPeriodNumber -gt 0) { [string]$summaryPreviousPeriodNumber } else { "" }
    Set-PeriodScopeLabelsForRows -Rows $databaseSquareRows -CurrentPeriodNumber $summaryCurrentPeriodNumber -PreviousPeriodNumber $summaryPreviousPeriodNumber
    Set-PeriodScopeLabelsForRows -Rows $missingTreeRows -CurrentPeriodNumber $summaryCurrentPeriodNumber -PreviousPeriodNumber $summaryPreviousPeriodNumber

    $databaseSquareReviewCount = Get-ReportReviewNeededCount -Rows $databaseSquareRows
    $databaseSquareNeedsReview = $databaseSquareReviewCount -gt 0
    $databaseSquarePeriodCounts = Get-SummaryPeriodErrorCounts -Rows $databaseSquareRows -RowKind "Report" -CurrentPeriodNumber $summaryCurrentPeriodNumber -PreviousPeriodNumber $summaryPreviousPeriodNumber
    $databaseSquareWorksheet = [pscustomobject]@{
        Name = $databaseSquareSheetName
        Headers = @("Check", "PeriodNumber", "PeriodScope", "ExpectedTable", "ExpectedCount", "ObservedTable", "ObservedCount", "Status", "Notes")
        Rows = $databaseSquareRows
        SummaryText = "What DADA checked: whether plot, tree, regen, and custom-measurement row counts line up by measurement period. Regen measurement and custom-measurement counts use the period segment in RegenMeasKey. Rows marked Review need hand review."
        TabColor = Get-WorksheetReviewTabColor -NeedsReview $databaseSquareNeedsReview
        ErrorCount = $databaseSquareReviewCount
        PeriodCounts = $databaseSquarePeriodCounts
        Tables = "Plots, PlotMeasurements, PlotCustomMeasurements, Trees, TreeMeasurements, TreeCustomMeasurements, RegenMeasurements, RegenCustomMeasurements"
        Rules = "Square count report"
    }
    [void]$projectCheckWorksheets.Add($databaseSquareWorksheet)

    $missingTreesReviewCount = Get-ReportReviewNeededCount -Rows $missingTreeRows
    $missingTreesNeedsReview = $missingTreesReviewCount -gt 0
    $missingTreesPeriodCounts = Get-SummaryPeriodErrorCounts -Rows $missingTreeRows -RowKind "Report" -CurrentPeriodNumber $summaryCurrentPeriodNumber -PreviousPeriodNumber $summaryPreviousPeriodNumber
    $missingTreesWorksheet = [pscustomobject]@{
        Name = $missingTreesSheetName
        Headers = @("Status", "PlotNumber", "TreeNumber", "SpeciesCode", "MinorPlot", "TreeID", "PeriodNumber", "PeriodScope", "RecordLabel", "Notes")
        Rows = $missingTreeRows
        SummaryText = "What DADA checked: trees that have at least one measurement were checked for missing TreeMeasurements rows in the selected measurement period scope. TreeHistory transition checks still review all periods separately."
        TabColor = Get-WorksheetReviewTabColor -NeedsReview $missingTreesNeedsReview
        ErrorCount = $missingTreesReviewCount
        PeriodCounts = $missingTreesPeriodCounts
        Tables = "Trees, TreeMeasurements, ProjectMeasurementPeriods"
        Rules = "Missing tree measurement report"
    }
    [void]$projectCheckWorksheets.Add($missingTreesWorksheet)

    $groups = @($exportRows.Rows | Group-Object { Get-ExportWorksheetGroupName -Row $_ } | Sort-Object Name)
    foreach ($group in $groups) {
        $tablesText = [string]::Join(", ", [string[]]@($group.Group | ForEach-Object { [string]$_["TableName"] } | Sort-Object -Unique))
        $rulesText = [string]::Join(", ", [string[]]@($group.Group | ForEach-Object { [string]$_["RuleName"] } | Sort-Object -Unique))
        $groupRows = @($group.Group)
        $findingCount = Get-ExportFindingCount -Rows ([object[]]$groupRows)
        $needsReview = $findingCount -gt 0
        $periodCounts = Get-SummaryPeriodErrorCounts -Rows ([object[]]$groupRows) -RowKind "Export" -CurrentPeriodNumber $summaryCurrentPeriodNumber -PreviousPeriodNumber $summaryPreviousPeriodNumber
        $worksheet = [pscustomobject]@{
            Name = (Get-SafeWorksheetName -Name $group.Name -UsedNames $usedSheetNames)
            Headers = (Get-ExportWorksheetHeaders -Rows ([object[]]$groupRows))
            Rows = $groupRows
            SummaryText = (Get-ExportWorksheetSummaryText -WorksheetName $group.Name -Rows ([object[]]$groupRows) -FindingCount $findingCount)
            TabColor = Get-WorksheetReviewTabColor -NeedsReview $needsReview
            ErrorCount = $findingCount
            PeriodCounts = $periodCounts
            Tables = $tablesText
            Rules = $rulesText
        }
        if ($needsReview) { [void]$reviewWorksheets.Add($worksheet) } else { [void]$passedWorksheets.Add($worksheet) }
    }

    $addSummaryWorksheetRow = {
        param(
            [object]$Worksheet,
            [bool]$SectionStart = $false
        )

        $periodCounts = [pscustomobject]@{ Current = 0; Previous = 0; Other = 0 }
        if ($null -ne $Worksheet -and $Worksheet.PSObject.Properties["PeriodCounts"] -and $null -ne $Worksheet.PeriodCounts) {
            $periodCounts = $Worksheet.PeriodCounts
        }
        [void]$summaryRows.Add([pscustomobject]@{
            Worksheet = $Worksheet.Name
            ErrorCount = $Worksheet.ErrorCount
            CurrentPeriodNumber = $summaryCurrentPeriodDisplay
            CurrentPeriodErrorCount = $periodCounts.Current
            PreviousPeriodNumber = $summaryPreviousPeriodDisplay
            PreviousPeriodErrorCount = $periodCounts.Previous
            OtherOrAllPeriodErrorCount = $periodCounts.Other
            Tables = $Worksheet.Tables
            Rules = $Worksheet.Rules
            SummarySectionStart = $SectionStart
        })
    }

    foreach ($worksheet in $projectCheckWorksheets) {
        & $addSummaryWorksheetRow $worksheet
    }
    [void]$summaryRows.Add([pscustomobject]@{
        Worksheet = $runDataSheetName
        ErrorCount = 0
        CurrentPeriodNumber = $summaryCurrentPeriodDisplay
        CurrentPeriodErrorCount = 0
        PreviousPeriodNumber = $summaryPreviousPeriodDisplay
        PreviousPeriodErrorCount = 0
        OtherOrAllPeriodErrorCount = 0
        Tables = "Run metadata"
        Rules = "Run settings, counts, and source context"
        SummarySectionStart = $false
    })
    $fieldSummaryWorksheets = @($reviewWorksheets.ToArray() + $passedWorksheets.ToArray())
    foreach ($sectionName in @("Plot data", "Tree data", "Regen data", "Other data")) {
        $sectionWorksheets = @(
            $fieldSummaryWorksheets |
                Where-Object { (Get-SummaryWorksheetSectionName -Worksheet $_) -eq $sectionName } |
                Sort-Object `
                    @{ Expression = { if ([int]$_.ErrorCount -gt 0) { 0 } else { 1 } } },
                    @{ Expression = { [string]$_.Name } }
        )
        $isFirstSectionRow = $true
        foreach ($worksheet in $sectionWorksheets) {
            & $addSummaryWorksheetRow $worksheet $isFirstSectionRow
            $isFirstSectionRow = $false
        }
    }

    [void]$worksheets.Add([pscustomobject]@{
        Name = $summarySheetName
        Headers = @("Worksheet", "ErrorCount", "CurrentPeriodNumber", "CurrentPeriodErrorCount", "PreviousPeriodNumber", "PreviousPeriodErrorCount", "OtherOrAllPeriodErrorCount", "Tables", "Rules")
        Rows = @($summaryRows.ToArray())
        SummaryText = "What DADA checked: this tab lists each worksheet, total records needing review, current-period and previous-period error counts, other/all-period/no-period counts, source tables, and represented rules."
        IsSummarySheet = $true
    })

    foreach ($worksheet in $projectCheckWorksheets) {
        [void]$worksheets.Add($worksheet)
    }

    [void]$worksheets.Add([pscustomobject]@{
        Name = $runDataSheetName
        Headers = @("Section", "Item", "Value")
        Rows = $runDataRows
        SummaryText = "What DADA checked: this tab records the run settings, source database, selected period scope, thresholds, AI/manual settings, source table counts, and audit counts used to create this workbook."
    })

    foreach ($worksheet in $reviewWorksheets) {
        [void]$worksheets.Add($worksheet)
    }
    foreach ($worksheet in $passedWorksheets) {
        [void]$worksheets.Add($worksheet)
    }
    $sheetBuildStopwatch.Stop()
    Add-GuideRunPerformanceEntry -Name "Export: build workbook sheet model" -Elapsed $sheetBuildStopwatch.Elapsed -CountText "$($worksheets.Count) worksheet(s)"

    $outputDirectory = [System.IO.Path]::GetDirectoryName($Path)
    if ([string]::IsNullOrWhiteSpace($outputDirectory)) { $outputDirectory = (Get-Location).Path }
    [System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
    $tempExportDirectory = Get-AppTempDirectory
    [System.IO.Directory]::CreateDirectory($tempExportDirectory) | Out-Null
    $tempRoot = Join-Path $tempExportDirectory ("~audit_export_" + [guid]::NewGuid().ToString("N"))

    try {
        $workbookWriteStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        [System.IO.Directory]::CreateDirectory($tempRoot) | Out-Null
        [System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "_rels")) | Out-Null
        [System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "docProps")) | Out-Null
        [System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "xl")) | Out-Null
        [System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "xl\_rels")) | Out-Null
        [System.IO.Directory]::CreateDirectory((Join-Path $tempRoot "xl\worksheets")) | Out-Null

        $contentTypes = New-Object System.Text.StringBuilder
        [void]$contentTypes.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$contentTypes.AppendLine('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
        [void]$contentTypes.AppendLine('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        [void]$contentTypes.AppendLine('<Default Extension="xml" ContentType="application/xml"/>')
        [void]$contentTypes.AppendLine('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
        [void]$contentTypes.AppendLine('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
        for ($i = 1; $i -le $worksheets.Count; $i++) {
            [void]$contentTypes.AppendLine("<Override PartName=""/xl/worksheets/sheet$i.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>")
        }
        [void]$contentTypes.AppendLine('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
        [void]$contentTypes.AppendLine('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
        [void]$contentTypes.AppendLine('</Types>')
        Write-Utf8File -Path (Join-Path $tempRoot "[Content_Types].xml") -Text $contentTypes.ToString()

        Write-Utf8File -Path (Join-Path $tempRoot "_rels\.rels") -Text @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

        Write-Utf8File -Path (Join-Path $tempRoot "docProps\core.xml") -Text ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:title>DADA Audit Export</dc:title><dc:creator>DADA - Database Dad</dc:creator><cp:lastModifiedBy>DADA - Database Dad</cp:lastModifiedBy><dcterms:created xsi:type=""dcterms:W3CDTF"">" + (Get-Date).ToUniversalTime().ToString("s") + "Z</dcterms:created></cp:coreProperties>")
        Write-Utf8File -Path (Join-Path $tempRoot "docProps\app.xml") -Text '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>DADA - Database Dad</Application></Properties>'
        Write-Utf8File -Path (Join-Path $tempRoot "xl\styles.xml") -Text (New-WorkbookStylesXml)

        $workbook = New-Object System.Text.StringBuilder
        [void]$workbook.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$workbook.AppendLine('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')
        for ($i = 0; $i -lt $worksheets.Count; $i++) {
            $sheetId = $i + 1
            $sheetName = ConvertTo-XmlText $worksheets[$i].Name
            [void]$workbook.AppendLine("<sheet name=""$sheetName"" sheetId=""$sheetId"" r:id=""rId$sheetId""/>")
        }
        [void]$workbook.AppendLine('</sheets></workbook>')
        Write-Utf8File -Path (Join-Path $tempRoot "xl\workbook.xml") -Text $workbook.ToString()

        $workbookRels = New-Object System.Text.StringBuilder
        [void]$workbookRels.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$workbookRels.AppendLine('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
        for ($i = 1; $i -le $worksheets.Count; $i++) {
            [void]$workbookRels.AppendLine("<Relationship Id=""rId$i"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$i.xml""/>")
        }
        [void]$workbookRels.AppendLine("<Relationship Id=""rIdStyles"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>")
        [void]$workbookRels.AppendLine('</Relationships>')
        Write-Utf8File -Path (Join-Path $tempRoot "xl\_rels\workbook.xml.rels") -Text $workbookRels.ToString()

        $runDataWorksheetIndex = -1
        $worksheetWriteOrder = New-Object System.Collections.Generic.List[int]
        for ($i = 0; $i -lt $worksheets.Count; $i++) {
            if ([string]$worksheets[$i].Name -eq $runDataSheetName) {
                $runDataWorksheetIndex = $i
            }
            else {
                [void]$worksheetWriteOrder.Add($i)
            }
        }
        if ($runDataWorksheetIndex -ge 0) {
            [void]$worksheetWriteOrder.Add($runDataWorksheetIndex)
        }

        $worksheetWriteNumber = 0
        foreach ($worksheetIndex in [int[]]$worksheetWriteOrder.ToArray()) {
            Assert-GuideRunNotCanceled
            $worksheetWriteNumber++
            if ($worksheetIndex -eq $runDataWorksheetIndex) {
                Set-RunDataRowValue -Rows $runDataRows -Section "Run" -Item "App run time" -Value (Get-CurrentGuideRunElapsedText)
            }

            $sheetPercent = $ProgressStart + [int]($progressSpan * (0.82 + (0.10 * ($worksheetWriteNumber / [double]$worksheets.Count))))
            Set-AppProgress -Status "Writing worksheet $worksheetWriteNumber of $($worksheets.Count)..." -Percent $sheetPercent
            $sheetPath = Join-Path $tempRoot ("xl\worksheets\sheet" + ($worksheetIndex + 1) + ".xml")
            $tabColor = ""
            if ($worksheets[$worksheetIndex].PSObject.Properties["TabColor"]) {
                $tabColor = [string]$worksheets[$worksheetIndex].TabColor
            }
            $summaryText = ""
            if ($worksheets[$worksheetIndex].PSObject.Properties["SummaryText"]) {
                $summaryText = [string]$worksheets[$worksheetIndex].SummaryText
            }
            $isSummarySheet = $false
            if ($worksheets[$worksheetIndex].PSObject.Properties["IsSummarySheet"]) {
                $isSummarySheet = [bool]$worksheets[$worksheetIndex].IsSummarySheet
            }
            $sheetXml = New-WorksheetXml -Headers ([string[]]$worksheets[$worksheetIndex].Headers) -Rows ([object[]]$worksheets[$worksheetIndex].Rows) -TabColor $tabColor -SummaryText $summaryText -IsSummarySheet:($isSummarySheet)
            Write-Utf8File -Path $sheetPath -Text $sheetXml
        }
        $workbookWriteStopwatch.Stop()
        Add-GuideRunPerformanceEntry -Name "Export: write worksheet XML files" -Elapsed $workbookWriteStopwatch.Elapsed -CountText "$($worksheets.Count) worksheet(s)"

        Assert-GuideRunNotCanceled
        Set-AppProgress -Status "Saving Excel workbook..." -Percent ($ProgressEnd - 2)
        $zipStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        if (Test-Path -LiteralPath $Path) {
            [System.IO.File]::Delete($Path)
        }
        $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            $files = Get-ChildItem -LiteralPath $tempRoot -Recurse -File
            foreach ($file in $files) {
                $entryName = $file.FullName.Substring($tempRoot.Length + 1).Replace("\", "/")
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, $entryName) | Out-Null
            }
        }
        finally {
            $archive.Dispose()
        }
        $zipStopwatch.Stop()
        Add-GuideRunPerformanceEntry -Name "Export: zip workbook package" -Elapsed $zipStopwatch.Elapsed -CountText "$($files.Count) file(s)"
    }
    finally {
        Clear-DirectoryTreeQuietly -Path $tempRoot
    }

    Set-AppProgress -Status "Excel workbook saved." -Percent $ProgressEnd
    return $exportRows.Rows.Count
}

function Export-AuditToExcel {
    if ([string]::IsNullOrWhiteSpace($script:DbPath) -or [string]::IsNullOrWhiteSpace($script:ConnectionString)) {
        [System.Windows.Forms.MessageBox]::Show("Connect to a database first.", $script:AppName) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Excel workbook (*.xlsx)|*.xlsx"
    $dialog.Title = "Save audit export"
    $dialog.FileName = "InventoryCleanAudit_" + (Get-Date -Format "yyyyMMdd_HHmm") + ".xlsx"
    if (-not [string]::IsNullOrWhiteSpace($script:DbPath)) {
        $dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:DbPath)
    }

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        $dialog.Dispose()
        return
    }

    $exportPath = $dialog.FileName
    $runLogPath = Get-GuideRunLogPathFromExportPath -ExportPath $exportPath
    $dialog.Dispose()

    $connection = $null
    try {
        Set-AppProgress -Status "Opening database for export..." -Percent 5
        $connection = Open-AccessConnection
        $useAiProjectGuidance = Test-AiProjectGuidanceSelected
        if ($useAiProjectGuidance) {
            Add-Log "AI guidance is on. The export will ask for AIMessage guidance using coded rules, template metadata, and manual excerpts when loaded."
            Add-Log "Compact AI export prompts: $(if (Test-AiCompactExportSelected) { 'On' } else { 'Off' })"
        }
        $count = Export-AuditWorkbook `
            -Connection $connection `
            -Path $exportPath `
            -UseAiProjectGuidance:$useAiProjectGuidance
        Add-Log "Exported $count audit rows to $exportPath"
        [System.Windows.Forms.MessageBox]::Show("Audit export complete.", $script:AppName) | Out-Null
    }
    catch {
        Set-AppProgress -Status "Export failed." -Percent 0
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Export failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        Add-Log "Export failed: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $connection) {
            $connection.Close()
            $connection.Dispose()
        }
    }
}

function Invoke-CfiGuideChecks {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $script:PeriodScopeConditionCache = @{}
    $script:AppColumnReviewMetadataCache = @{}
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Preparing audit tables..." -Percent 10
    [void](Invoke-GuideTimedStep -Name "Prepare audit table" -ScriptBlock {
        Ensure-AuditTable -Connection $Connection
        [void](Invoke-NonQuery -Connection $Connection -Sql "DELETE FROM [InventoryCleanAudit]")
    })

    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Checking workbook-style counts..." -Percent 20
    $countAuditCount = [int](Invoke-GuideTimedStep -Name "Workbook-style count checks" -ScriptBlock {
        Add-WorkbookCountChecks -Connection $Connection
    })
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Running reference checks..." -Percent 32
    $rmdAuditCount = [int](Invoke-GuideTimedStep -Name "R-script reference checks" -ScriptBlock {
        Add-RmdReferenceChecks -Connection $Connection
    })
    $codeAuditCount = 0
    $rangeAuditCount = 0

    $targetTables = @(Invoke-GuideTimedStep -Name "Load active workbook target tables" -ScriptBlock {
        Get-CfiWorkbookTargetTables -Connection $Connection
    })
    $tableIndex = 0
    foreach ($tableName in $targetTables) {
        Assert-GuideRunNotCanceled
        $tableIndex++
        $tablePercent = 42
        if ($targetTables.Count -gt 0) {
            $tablePercent = 42 + [int](($tableIndex / [double]$targetTables.Count) * 28)
        }
        Set-AppProgress -Status "Checking $tableName ($tableIndex of $($targetTables.Count))..." -Percent $tablePercent
        [void](Invoke-GuideTimedStep -Name "Prepare table metadata: $tableName" -ScriptBlock {
            Ensure-NeedsReviewColumn -Connection $Connection -TableName $tableName
        })
        $columns = @(Invoke-GuideTimedStep -Name "Read columns: $tableName" -ScriptBlock {
            Get-TableColumns -Connection $Connection -TableName $tableName
        })
        $codeAuditCount += [int](Invoke-GuideTimedStep -Name "Code validation: $tableName" -ScriptBlock {
            Add-CodeValidationAudit -Connection $Connection -Transaction $null -TableName $tableName -Columns $columns
        })
        if (Test-RangeChecksEnabled) {
            $rangeAuditCount += [int](Invoke-GuideTimedStep -Name "Range and required-field checks: $tableName" -ScriptBlock {
                Add-CfiRangeChecksForTable -Connection $Connection -TableName $tableName -Columns $columns
            })
        }
    }

    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Checking PlotStatus progression..." -Percent 70
    $plotStatusProgressionAuditCount = [int](Invoke-GuideTimedStep -Name "PlotStatus progression checks" -ScriptBlock {
        Add-PlotStatusProgressionChecks -Connection $Connection
    })
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Checking plot data on not-measured/dropped plots..." -Percent 71
    $inactivePlotDataAuditCount = [int](Invoke-GuideTimedStep -Name "Plot data on not-measured/dropped plots" -ScriptBlock {
        Add-NoPlotDataOnInactivePlotStatusChecks -Connection $Connection
    })
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Checking tree data on not-measured/dropped plots..." -Percent 72
    $inactivePlotTreeDataAuditCount = [int](Invoke-GuideTimedStep -Name "Tree data on not-measured/dropped plots" -ScriptBlock {
        Add-NoTreeDataOnInactivePlotStatusChecks -Connection $Connection
    })
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Checking regen data on not-measured/dropped plots..." -Percent 73
    $inactivePlotRegenDataAuditCount = [int](Invoke-GuideTimedStep -Name "Regen data on not-measured/dropped plots" -ScriptBlock {
        Add-NoRegenDataOnInactivePlotStatusChecks -Connection $Connection
    })
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Checking tree remeasurements..." -Percent 74
    $remeasurementAuditCount = [int](Invoke-GuideTimedStep -Name "Tree remeasurement checks" -ScriptBlock {
        Add-TreeRemeasurementChecks -Connection $Connection
    })
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Checking TreeClass/TreeStatus conversions..." -Percent 76
    $treeHistoryConversionAuditCount = [int](Invoke-GuideTimedStep -Name "TreeClass/TreeStatus conversion checks" -ScriptBlock {
        Add-TreeHistoryConversionChecks -Connection $Connection
    })
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Checking ingrowth TreeHistory rules..." -Percent 77
    $ingrowthAuditCount = [int](Invoke-GuideTimedStep -Name "Ingrowth TreeHistory checks" -ScriptBlock {
        Add-IngrowthTreeHistoryChecks -Connection $Connection
    })
    Assert-GuideRunNotCanceled
    Set-AppProgress -Status "Verifying woodland species DBH shrink exclusions..." -Percent 78
    $woodlandShrinkSkipVerification = Invoke-GuideTimedStep -Name "Woodland DBH shrink exclusion verification" -ScriptBlock {
        Get-WoodlandShrinkSkipVerification -Connection $Connection
    }

    return [pscustomobject]@{
        WorkbookCounts = $countAuditCount
        RScriptReference = $rmdAuditCount
        InvalidCodes = $codeAuditCount
        Ranges = $rangeAuditCount
        PlotStatusProgression = $plotStatusProgressionAuditCount
        InactivePlotData = $inactivePlotDataAuditCount
        InactivePlotTreeData = $inactivePlotTreeDataAuditCount
        InactivePlotRegenData = $inactivePlotRegenDataAuditCount
        Remeasurement = $remeasurementAuditCount
        TreeHistoryConversion = $treeHistoryConversionAuditCount
        Ingrowth = $ingrowthAuditCount
        WoodlandShrinkSkipVerification = $woodlandShrinkSkipVerification
    }
}

function Invoke-GuideChecksToWorkbook {
    param(
        [string]$SourcePath,
        [string]$Password,
        [string]$ExportPath,
        [switch]$UseAiProjectGuidance
    )

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        throw "The selected database file does not exist."
    }

    $tempDirectory = Get-AppTempDirectory
    [System.IO.Directory]::CreateDirectory($tempDirectory) | Out-Null
    $extension = [System.IO.Path]::GetExtension($SourcePath)
    $tempPath = Join-Path $tempDirectory ("guide_checks_" + [guid]::NewGuid().ToString("N") + $extension)
    $lockPath = [System.IO.Path]::ChangeExtension($tempPath, ".ldb")
    if ($extension -and $extension.ToLowerInvariant() -eq ".accdb") {
        $lockPath = [System.IO.Path]::ChangeExtension($tempPath, ".laccdb")
    }

    $connection = $null
    Initialize-GuideRunPerformance
    try {
        try {
            $sourceItem = Get-Item -LiteralPath $SourcePath
            $tempRoot = [System.IO.Path]::GetPathRoot($tempDirectory)
            $driveInfo = if (-not [string]::IsNullOrWhiteSpace($tempRoot)) { New-Object System.IO.DriveInfo -ArgumentList $tempRoot } else { $null }
            $freeText = if ($null -ne $driveInfo) { Format-ByteSizeText -Bytes $driveInfo.AvailableFreeSpace } else { "not available" }
            Write-GuideRunLogEntry -Message "Performance context: source database size $(Format-ByteSizeText -Bytes $sourceItem.Length); temp folder $tempDirectory; temp drive free space $freeText."
        }
        catch {
            Write-GuideRunLogEntry -Message "Performance context: could not read source/temp disk details. $($_.Exception.Message)"
        }

        Assert-GuideRunNotCanceled
        Set-AppProgress -Status "Copying database to a temporary file..." -Percent 5
        [void](Invoke-GuideTimedStep -Name "Copy database to temporary file" -ScriptBlock {
            Copy-Item -LiteralPath $SourcePath -Destination $tempPath -Force
        })
        Assert-GuideRunNotCanceled
        Set-AppProgress -Status "Opening temporary database..." -Percent 8
        $connection = Invoke-GuideTimedStep -Name "Open temporary Access database" -ScriptBlock {
            $tempConnectionString = New-AccessConnectionString -Path $tempPath -Password $Password
            $openedConnection = New-Object System.Data.OleDb.OleDbConnection($tempConnectionString)
            $openedConnection.Open()
            return $openedConnection
        }
        Assert-GuideRunNotCanceled
        $counts = Invoke-GuideTimedStep -Name "Run all cleaning checks" -ScriptBlock {
            Invoke-CfiGuideChecks -Connection $connection
        }
        Assert-GuideRunNotCanceled
        Set-AppProgress -Status "Exporting workbook..." -Percent 78
        $exportedRows = [int](Invoke-GuideTimedStep -Name "Export Excel workbook" -ScriptBlock {
            Export-AuditWorkbook `
                -Connection $connection `
                -Path $ExportPath `
                -UseAiProjectGuidance:$UseAiProjectGuidance `
                -RunCounts $counts `
                -ProgressStart 78 `
                -ProgressEnd 98
        })
        $counts | Add-Member -MemberType NoteProperty -Name ExportedRows -Value $exportedRows
        $counts | Add-Member -MemberType NoteProperty -Name ReviewedCleanRows -Value ([int]$script:LastReviewedCleanExportRows)
        $counts | Add-Member -MemberType NoteProperty -Name AiSuggestedEditGroupsApplied -Value ([int]$script:LastAiSuggestedEditGroupsApplied)
        $counts | Add-Member -MemberType NoteProperty -Name AiSuggestedEditGroupsSkipped -Value ([int]$script:LastAiSuggestedEditGroupsSkipped)
        $counts | Add-Member -MemberType NoteProperty -Name AiSuggestedEditRowsApplied -Value ([int]$script:LastAiSuggestedEditRowsApplied)
        $counts | Add-Member -MemberType NoteProperty -Name AiSuggestedEditRowsSkipped -Value ([int]$script:LastAiSuggestedEditRowsSkipped)
        $counts | Add-Member -MemberType NoteProperty -Name AiManualGroupsUsed -Value ([int]$script:LastAiSuggestedEditManualGroupsUsed)
        $counts | Add-Member -MemberType NoteProperty -Name AiManualRowsUsed -Value ([int]$script:LastAiSuggestedEditManualRowsUsed)
        $counts | Add-Member -MemberType NoteProperty -Name AiNoManualRetryGroups -Value ([int]$script:LastAiSuggestedEditNoManualRetryGroups)
        $counts | Add-Member -MemberType NoteProperty -Name AiNoManualRetryRows -Value ([int]$script:LastAiSuggestedEditNoManualRetryRows)
        $counts | Add-Member -MemberType NoteProperty -Name AiNoManualSnippetGroups -Value ([int]$script:LastAiSuggestedEditNoManualSnippetGroups)
        $counts | Add-Member -MemberType NoteProperty -Name AiNoManualSnippetRows -Value ([int]$script:LastAiSuggestedEditNoManualSnippetRows)
        Write-GuideRunPerformanceSummary -Status "Completed"
        Set-AppProgress -Status "Run complete." -Percent 100
        return $counts
    }
    catch [System.OperationCanceledException] {
        Write-GuideRunPerformanceSummary -Status "Canceled"
        throw
    }
    catch {
        Write-GuideRunPerformanceSummary -Status "Failed"
        throw
    }
    finally {
        if ($null -ne $connection) {
            $connection.Close()
            $connection.Dispose()
        }
        foreach ($pathToDelete in @($lockPath, $tempPath)) {
            if (-not [string]::IsNullOrWhiteSpace($pathToDelete) -and (Test-Path -LiteralPath $pathToDelete)) {
                try { [System.IO.File]::Delete($pathToDelete) } catch { }
            }
        }
    }
}

function Format-ElapsedText {
    param([TimeSpan]$Elapsed)

    if ($Elapsed.TotalHours -ge 1) {
        return "{0:00}:{1:00}:{2:00}" -f [Math]::Floor($Elapsed.TotalHours), $Elapsed.Minutes, $Elapsed.Seconds
    }

    return "{0:00}:{1:00}" -f [Math]::Floor($Elapsed.TotalMinutes), $Elapsed.Seconds
}

function ConvertTo-GuideRunDateTime {
    param([object]$Value)

    if ($Value -is [datetime]) { return [datetime]$Value }
    if ($null -eq $Value) { return Get-Date }

    $text = ""
    try {
        if ($Value.PSObject.Properties["value"]) {
            $text = [string]$Value.value
        }
        elseif ($Value.PSObject.Properties["DateTime"]) {
            $text = [string]$Value.DateTime
        }
        else {
            $text = [string]$Value
        }
    }
    catch {
        $text = [string]$Value
    }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
        return $parsed
    }
    if ([datetime]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    return Get-Date
}

function Set-ControlEnabledQuietly {
    param(
        [object]$Control,
        [bool]$Enabled
    )

    try {
        if ($null -ne $Control) { $Control.Enabled = $Enabled }
    }
    catch {
    }
}

function Set-GuideRunControlsState {
    param([bool]$Running)

    $enabled = -not $Running
    foreach ($control in @(
        $browseButton,
        $clearButton,
        $projectManualButton,
        $pathBox,
        $passwordBox,
        $tableCombo,
        $topGuideButton,
        $workbookChecksButton,
        $cleanButton,
        $previewButton,
        $aiEnabledCheck,
        $topAiEnabledCheck,
        $aiHelperButton,
        $topAiHelperButton,
        $aiExportGuidanceCheck,
        $aiCompactExportCheck,
        $splitHistoryColumnsCheck,
        $trimTextCheck,
        $blankNullCheck,
        $speciesUpperCheck,
        $rangeCheck,
        $duplicateCheck,
        $dbhMax,
        $heightMax,
        $dbhGrowthMax,
        $heightGrowthMax,
        $periodScopeCheck,
        $singleMeasurementProjectCheck,
        $currentPeriodBox,
        $pastPeriodBox,
        $stemCountMax,
        $treeStemCountCheck,
        $newMortalityIdbhCheck,
        $oldMortalityIdbhCheck,
        $newMortalityHeightCheck,
        $oldMortalityHeightCheck,
        $problem127HeightCheck,
        $problem128HeightCheck,
        $problem123HeightCheck,
        $heightProtocolCombo,
        $heightSubsampleMinimumBox,
        $heightSubsampleMinIdbhBox,
        $heightSubsampleAllAtOrAboveCheck,
        $heightSubsampleAllAtOrAboveBox,
        $heightRareSpeciesCodesBox,
        $heightRequiredMinorPlotsBox,
        $regenTimberSeedlingMinorPlotsBox,
        $regenTimberSapling20MinorPlotsBox,
        $regenTimberSapling40MinorPlotsBox,
        $regenWoodlandSeedlingMinorPlotsBox,
        $regenWoodlandSapling20MinorPlotsBox,
        $regenWoodlandSapling40MinorPlotsBox,
        $woodlandSpeciesCodesBox
    )) {
        Set-ControlEnabledQuietly -Control $control -Enabled $enabled
    }

    if ($enabled) {
        try {
            $periodFieldsEnabled = ($null -ne $periodScopeCheck -and $periodScopeCheck.Checked)
            $pastPeriodEnabled = ($periodFieldsEnabled -and -not (Test-SingleMeasurementProjectEnabled))
            Set-ControlEnabledQuietly -Control $singleMeasurementProjectCheck -Enabled $periodFieldsEnabled
            Set-ControlEnabledQuietly -Control $currentPeriodBox -Enabled $periodFieldsEnabled
            Set-ControlEnabledQuietly -Control $pastPeriodBox -Enabled $pastPeriodEnabled
            Update-HeightProtocolControlState
        }
        catch {
        }
    }

    foreach ($cancelButton in @($cancelRunButton)) {
        try {
            if ($null -ne $cancelButton) {
                $cancelButton.Visible = $Running
                $cancelButton.Enabled = $Running
            }
        }
        catch {
        }
    }
}

function Get-GuideRunSettingsSnapshot {
    param(
        [string]$ExportPath,
        [switch]$ForceDisableAiProjectGuidance
    )

    $useAiProjectGuidance = if ($ForceDisableAiProjectGuidance) { $false } else { Test-AiProjectGuidanceSelected }
    $manualText = $script:ProjectManualText
    $manualLoadedFrom = $script:ProjectManualLoadedFrom
    $periodScopeEnabled = [bool]$periodScopeCheck.Checked
    $singleMeasurementProject = [bool]$singleMeasurementProjectCheck.Checked
    if ($singleMeasurementProject) {
        $currentPeriodNumber = 1
        $pastPeriodNumber = 0
    }
    else {
        $currentPeriodNumber = [int]$currentPeriodBox.Value
        $pastPeriodNumber = [int]$pastPeriodBox.Value
    }

    if ($periodScopeEnabled -and $currentPeriodNumber -le 0) {
        throw "Enter the current measurement period number before running, or turn off 'Limit normal checks to measurement period(s)'."
    }
    if ($periodScopeEnabled -and (-not $singleMeasurementProject) -and $pastPeriodNumber -le 0) {
        throw "Enter the previous measurement period number, or check 'Single measurement project' when there is no previous period."
    }

    if ($useAiProjectGuidance -and -not [string]::IsNullOrWhiteSpace($script:ProjectManualPath) -and [string]::IsNullOrWhiteSpace($manualText)) {
        Set-ProjectManualFile -Path $script:ProjectManualPath
        $manualText = $script:ProjectManualText
        $manualLoadedFrom = $script:ProjectManualLoadedFrom
    }

    $settings = [pscustomobject]@{
        SourcePath = [string]$script:DbPath
        Password = [string]$passwordBox.Text
        ExportPath = [string]$ExportPath
        RunLogPath = [string](Get-GuideRunLogPathFromExportPath -ExportPath $ExportPath)
        UseAiProjectGuidance = [bool]$useAiProjectGuidance
        AiCompactExport = [bool](Test-AiCompactExportSelected)
        SplitHistoryColumns = $true
        AiHelperOn = [bool](Test-AiHelperIsOn)
        AiEndpoint = [string]$script:AiDefaultEndpoint
        AiModel = [string]$script:AiDefaultModel
        AiApiKey = [string]$script:AiApiKey
        ProjectManualPath = [string]$script:ProjectManualPath
        ProjectManualText = [string]$manualText
        ProjectManualLoadedFrom = [string]$manualLoadedFrom
        RangeChecksEnabled = [bool]$rangeCheck.Checked
        TreeStemCountChecksEnabled = [bool]$treeStemCountCheck.Checked
        PeriodScopeEnabled = [bool]$periodScopeEnabled
        SingleMeasurementProject = [bool]$singleMeasurementProject
        CurrentPeriodNumber = [int]$currentPeriodNumber
        PastPeriodNumber = [int]$pastPeriodNumber
        NewMortalityIdbhChecksEnabled = [bool]$newMortalityIdbhCheck.Checked
        OldMortalityIdbhChecksEnabled = [bool]$oldMortalityIdbhCheck.Checked
        NewMortalityHeightChecksEnabled = [bool]$newMortalityHeightCheck.Checked
        OldMortalityHeightChecksEnabled = [bool]$oldMortalityHeightCheck.Checked
        Problem127HeightChecksEnabled = [bool]$problem127HeightCheck.Checked
        Problem128HeightChecksEnabled = [bool]$problem128HeightCheck.Checked
        Problem123HeightChecksEnabled = [bool]$problem123HeightCheck.Checked
        TotalHeightProtocolMode = [string](Get-TotalHeightProtocolMode)
        HeightSubsampleMinimumCount = [decimal]$heightSubsampleMinimumBox.Value
        HeightSubsampleMinimumIdbh = [decimal]$heightSubsampleMinIdbhBox.Value
        HeightSubsampleAllAtOrAboveEnabled = [bool]$heightSubsampleAllAtOrAboveCheck.Checked
        HeightSubsampleAllAtOrAboveIdbh = [decimal]$heightSubsampleAllAtOrAboveBox.Value
        HeightRareSpeciesCodesText = [string]$heightRareSpeciesCodesBox.Text
        DbhMax = [decimal]$dbhMax.Value
        HeightMax = [decimal]$heightMax.Value
        StemCountMax = [decimal]$stemCountMax.Value
        DbhGrowthMax = [decimal]$dbhGrowthMax.Value
        HeightGrowthMax = [decimal]$heightGrowthMax.Value
        WoodlandSpeciesCodesText = [string]$woodlandSpeciesCodesBox.Text
        HeightRequiredMinorPlotsText = [string]$heightRequiredMinorPlotsBox.Text
        RegenTimberSeedlingMinorPlotsText = [string]$regenTimberSeedlingMinorPlotsBox.Text
        RegenTimberSapling20MinorPlotsText = [string]$regenTimberSapling20MinorPlotsBox.Text
        RegenTimberSapling40MinorPlotsText = [string]$regenTimberSapling40MinorPlotsBox.Text
        RegenWoodlandSeedlingMinorPlotsText = [string]$regenWoodlandSeedlingMinorPlotsBox.Text
        RegenWoodlandSapling20MinorPlotsText = [string]$regenWoodlandSapling20MinorPlotsBox.Text
        RegenWoodlandSapling40MinorPlotsText = [string]$regenWoodlandSapling40MinorPlotsBox.Text
        StartedAt = (Get-Date).ToString("o")
    }

    if ($settings.UseAiProjectGuidance) {
        if (-not $settings.AiHelperOn) { throw "Turn on AI helper before using AI guidance in the export." }
        [void](Get-ValidatedAiEndpoint -Endpoint $settings.AiEndpoint)
        if ([string]::IsNullOrWhiteSpace($settings.AiModel)) { throw "Open AI helper and enter a model." }
        if ([string]::IsNullOrWhiteSpace($settings.AiApiKey)) { throw "Open AI helper and enter an API key. The key is only kept while this app window is open." }
    }

    return $settings
}

function Test-GuideRunActive {
    try {
        if ($null -eq $script:GuideRunWorker) { return $false }
        if ($script:GuideRunWorker.PSObject.Properties["Completed"] -and [bool]$script:GuideRunWorker.Completed) { return $false }
        return $true
    }
    catch {
        return $false
    }
}

function Request-GuideRunCancel {
    if (-not (Test-GuideRunActive)) { return }

    $script:GuideRunCancellationRequested = $true
    try {
        if (-not [string]::IsNullOrWhiteSpace($script:GuideRunCancelPath)) {
            [System.IO.File]::WriteAllText($script:GuideRunCancelPath, (Get-Date).ToString("o"), [System.Text.Encoding]::UTF8)
        }
    }
    catch {
    }

    foreach ($cancelButton in @($cancelRunButton, $topCancelRunButton)) {
        try {
            if ($null -ne $cancelButton) { $cancelButton.Enabled = $false }
        }
        catch {
        }
    }
    Set-AppProgress -Status "Cancel requested. Finishing the current step..." -Indeterminate
    Add-Log "Cancel requested. The app will stop after the current query or AI request finishes."
}

function Confirm-AppCloseWhileRunActive {
    param([System.Windows.Forms.FormClosingEventArgs]$EventArgs)

    if (-not (Test-GuideRunActive)) { return }

    if ($script:CloseAfterGuideRunCancel) {
        $EventArgs.Cancel = $true
        return
    }

    $message = "DADA is still processing a run.`r`n`r`nClosing now could leave the export unfinished. Do you want to request cancel and close after the current step finishes?"
    $choice = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "Run still processing",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        $EventArgs.Cancel = $true
        $script:CloseAfterGuideRunCancel = $true
        $script:SuppressGuideRunCompletionDialogs = $true
        Request-GuideRunCancel
        Add-Log "DADA will close after the active run stops."
        Set-AppProgress -Status "Cancel requested. DADA will close after the current step..." -Indeterminate
    }
    else {
        $EventArgs.Cancel = $true
    }
}

function Start-GuideChecksBackgroundRun {
    param([object]$Settings)

    if (Test-GuideRunActive) {
        [System.Windows.Forms.MessageBox]::Show("A run is already active.", $script:AppName) | Out-Null
        return
    }

    $script:GuideRunSettings = $Settings
    $script:GuideRunStartedAt = ConvertTo-GuideRunDateTime $Settings.StartedAt
    $script:GuideRunCancellationRequested = $false
    $script:AiGuidanceCache = @{}
    Initialize-GuideRunLog -Settings $Settings

    $tempDirectory = Get-AppTempDirectory
    [System.IO.Directory]::CreateDirectory($tempDirectory) | Out-Null
    $runId = [guid]::NewGuid().ToString("N")
    $progressPath = Join-Path $tempDirectory "guide_run_progress_$runId.json"
    $resultPath = Join-Path $tempDirectory "guide_run_result_$runId.json"
    $cancelPath = Join-Path $tempDirectory "guide_run_cancel_$runId.flag"
    $script:GuideRunCancelPath = $cancelPath

    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.Open()
    $powershell = [System.Management.Automation.PowerShell]::Create()
    $powershell.Runspace = $runspace
    $workerScript = @'
param(
    [string]$ScriptPath,
    [object]$Settings,
    [string]$ProgressPath,
    [string]$CancelPath,
    [string]$ResultPath
)

$ErrorActionPreference = "Stop"
$env:FOREST_CLEANER_NO_UI = "1"
. $ScriptPath

function Write-GuideRunResult {
    param([object]$Result)
    $json = $Result | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($ResultPath, $json, [System.Text.Encoding]::UTF8)
}

try {
    $settings = $Settings
    $script:GuideRunSettings = $settings
    $script:GuideRunStartedAt = ConvertTo-GuideRunDateTime $settings.StartedAt
    $script:GuideRunProgressPath = $ProgressPath
    $script:GuideRunCancelPath = $CancelPath
    $script:GuideRunLogPath = [string]$settings.RunLogPath
    $script:GuideRunCancellationRequested = $false
    $script:AiDefaultEndpoint = [string]$settings.AiEndpoint
    $script:AiDefaultModel = [string]$settings.AiModel
    $script:AiApiKey = [string]$settings.AiApiKey
    $script:ProjectManualPath = [string]$settings.ProjectManualPath
    $script:ProjectManualText = [string]$settings.ProjectManualText
    $script:ProjectManualLoadedFrom = [string]$settings.ProjectManualLoadedFrom

    $counts = Invoke-GuideChecksToWorkbook `
        -SourcePath ([string]$settings.SourcePath) `
        -Password ([string]$settings.Password) `
        -ExportPath ([string]$settings.ExportPath) `
        -UseAiProjectGuidance:([bool]$settings.UseAiProjectGuidance)

    Assert-GuideRunNotCanceled
    Write-GuideRunResult ([pscustomobject]@{
        Status = "Success"
        Counts = $counts
        ExportPath = [string]$settings.ExportPath
        UseAiProjectGuidance = [bool]$settings.UseAiProjectGuidance
        AiCompactExport = [bool]$settings.AiCompactExport
        CompletedAt = (Get-Date).ToString("o")
    })
}
catch [System.OperationCanceledException] {
    Write-GuideRunResult ([pscustomobject]@{
        Status = "Canceled"
        Message = "Run canceled."
        CompletedAt = (Get-Date).ToString("o")
    })
}
catch {
    Write-GuideRunResult ([pscustomobject]@{
        Status = "Error"
        Message = $_.Exception.Message
        Detail = ($_ | Out-String)
        CompletedAt = (Get-Date).ToString("o")
    })
}
'@

    [void]$powershell.AddScript($workerScript)
    $scriptPathForWorker = $PSCommandPath
    if ([string]::IsNullOrWhiteSpace($scriptPathForWorker)) {
        $scriptPathForWorker = Join-Path $PSScriptRoot "ForestInventoryCleaner.ps1"
    }
    [void]$powershell.AddArgument($scriptPathForWorker)
    [void]$powershell.AddArgument($Settings)
    [void]$powershell.AddArgument($progressPath)
    [void]$powershell.AddArgument($cancelPath)
    [void]$powershell.AddArgument($resultPath)
    $asyncResult = $powershell.BeginInvoke()

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $state = [pscustomobject]@{
        PowerShell = $powershell
        Runspace = $runspace
        AsyncResult = $asyncResult
        Timer = $timer
        ProgressPath = $progressPath
        ResultPath = $resultPath
        CancelPath = $cancelPath
        Completed = $false
    }
    $script:GuideRunWorker = $state

    $timer.Add_Tick({
        $state = $script:GuideRunWorker
        if ($null -eq $state) { return }

        try {
            if (Test-Path -LiteralPath $state.ProgressPath) {
                $progress = Get-Content -LiteralPath $state.ProgressPath -Raw | ConvertFrom-Json
                Set-AppProgress `
                    -Status ([string]$progress.Status) `
                    -Percent ([int]$progress.Percent) `
                    -Indeterminate:([bool]$progress.Indeterminate) `
                    -Clear:([bool]$progress.Clear)
            }
        }
        catch {
        }

        if (-not $state.AsyncResult.IsCompleted) { return }

        $state.Timer.Stop()
        $state.Completed = $true

        $settings = $script:GuideRunSettings
        $startedAt = $script:GuideRunStartedAt
        if ($null -eq $startedAt -and $null -ne $settings) { $startedAt = ConvertTo-GuideRunDateTime $settings.StartedAt }
        if ($null -eq $startedAt) { $startedAt = Get-Date }
        $elapsed = (Get-Date) - $startedAt
        $elapsedText = Format-ElapsedText -Elapsed $elapsed

        try {
            $earlyFailureDetail = ""
            try { $state.PowerShell.EndInvoke($state.AsyncResult) } catch { $earlyFailureDetail = $_.Exception.Message }
            try {
                $streamErrors = @($state.PowerShell.Streams.Error | ForEach-Object { $_.ToString() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                if ($streamErrors.Count -gt 0) {
                    $streamErrorText = [string]::Join("`r`n", [string[]]$streamErrors)
                    if ([string]::IsNullOrWhiteSpace($earlyFailureDetail)) {
                        $earlyFailureDetail = $streamErrorText
                    }
                    else {
                        $earlyFailureDetail += "`r`n" + $streamErrorText
                    }
                }
            }
            catch {
            }
            $result = $null
            if (Test-Path -LiteralPath $state.ResultPath) {
                $result = Get-Content -LiteralPath $state.ResultPath -Raw | ConvertFrom-Json
            }

            Set-GuideRunControlsState -Running $false

            if ($null -eq $result) {
                Set-AppProgress -Status "Run failed after $elapsedText." -Percent 0
                if ([string]::IsNullOrWhiteSpace($earlyFailureDetail)) {
                    $earlyFailureDetail = "The background runner stopped before it could report the detailed error."
                }
                $fullEarlyFailureDetail = $earlyFailureDetail
                $earlyFailureDetail = Limit-TextLength -Text $earlyFailureDetail -MaxLength 1200
                Add-Log "Run failed before returning a result: $earlyFailureDetail"
                Write-GuideRunFailureLog -Status "Failed before result" -Message $earlyFailureDetail -Detail $fullEarlyFailureDetail
                if (-not [string]::IsNullOrWhiteSpace([string]$script:GuideRunLogPath)) {
                    Add-Log "Failure details saved to run log: $script:GuideRunLogPath"
                }
                $failureMessage = "Run failed before returning a result.`r`n`r`n$earlyFailureDetail"
                if (-not [string]::IsNullOrWhiteSpace([string]$script:GuideRunLogPath)) {
                    $failureMessage += "`r`n`r`nRun log:`r`n$script:GuideRunLogPath"
                }
                [System.Windows.Forms.MessageBox]::Show($failureMessage, "Run failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }

            switch ([string]$result.Status) {
                "Canceled" {
                    Set-AppProgress -Status "Run canceled after $elapsedText." -Percent 0
                    Add-Log "Run canceled after $elapsedText."
                    Write-GuideRunFailureLog -Status "Canceled" -Message "Run canceled by the user."
                    if (-not [string]::IsNullOrWhiteSpace([string]$script:GuideRunLogPath)) {
                        Add-Log "Run log saved: $script:GuideRunLogPath"
                    }
                    if (-not $script:SuppressGuideRunCompletionDialogs) {
                        [System.Windows.Forms.MessageBox]::Show("The run was canceled after $elapsedText. No completed workbook was saved.", "Run canceled") | Out-Null
                    }
                }
                "Error" {
                    Set-AppProgress -Status "Run failed after $elapsedText." -Percent 0
                    Add-Log "Run failed: $($result.Message)"
                    $errorDetail = if ($result.PSObject.Properties["Detail"]) { [string]$result.Detail } else { "" }
                    Write-GuideRunFailureLog -Status "Failed" -Message ([string]$result.Message) -Detail $errorDetail
                    if (-not [string]::IsNullOrWhiteSpace([string]$script:GuideRunLogPath)) {
                        Add-Log "Failure details saved to run log: $script:GuideRunLogPath"
                    }
                    if (-not $script:SuppressGuideRunCompletionDialogs) {
                        $failureMessage = [string]$result.Message
                        if (-not [string]::IsNullOrWhiteSpace([string]$script:GuideRunLogPath)) {
                            $failureMessage += "`r`n`r`nRun log:`r`n$script:GuideRunLogPath"
                        }
                        [System.Windows.Forms.MessageBox]::Show($failureMessage, "Run failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                    }
                }
                default {
                    $counts = $result.Counts
                    Set-AppProgress -Status "Run complete after $elapsedText." -Percent 100
                    Add-Log "Run complete. Excel saved: $($result.ExportPath)"
                    Add-Log "Workbook count audit records: $($counts.WorkbookCounts)"
                    Add-Log "R-script reference audit records: $($counts.RScriptReference)"
                    Add-Log "Invalid CFI code audit records: $($counts.InvalidCodes)"
                    Add-Log "IDBH/height/stem/tree species/crown/radial/required plot-check audit records: $($counts.Ranges)"
                    if ($counts.PSObject.Properties["PlotStatusProgression"]) {
                        Add-Log "PlotStatus progression review records: $($counts.PlotStatusProgression)"
                    }
                    if ($counts.PSObject.Properties["InactivePlotData"]) {
                        Add-Log "Plot data on not-measured/dropped plot records: $($counts.InactivePlotData)"
                    }
                    if ($counts.PSObject.Properties["InactivePlotTreeData"]) {
                        Add-Log "Tree data on not-measured/dropped plot records: $($counts.InactivePlotTreeData)"
                    }
                    if ($counts.PSObject.Properties["InactivePlotRegenData"]) {
                        Add-Log "Regen data on not-measured/dropped plot records: $($counts.InactivePlotRegenData)"
                    }
                    Add-Log "Remeasurement consistency audit records: $($counts.Remeasurement)"
                    if ($counts.PSObject.Properties["TreeHistoryConversion"]) {
                        Add-Log "TreeHistory conversion review records: $($counts.TreeHistoryConversion)"
                    }
                    if ($counts.PSObject.Properties["Ingrowth"]) {
                        Add-Log "Ingrowth TreeHistory review records: $($counts.Ingrowth)"
                    }
                    if ($counts.PSObject.Properties["WoodlandShrinkSkipVerification"]) {
                        $woodlandVerification = $counts.WoodlandShrinkSkipVerification
                        if ($null -ne $woodlandVerification -and $woodlandVerification.PSObject.Properties["Message"]) {
                            Add-Log "Woodland DBH shrink exclusion verification: $($woodlandVerification.Message)"
                        }
                    }
                    Add-Log "Excel rows exported: $($counts.ExportedRows)"
                    if ($counts.PSObject.Properties["ReviewedCleanRows"]) {
                        Add-Log "Reviewed-clean field rows added: $($counts.ReviewedCleanRows)"
                    }
                    if ([bool]$result.UseAiProjectGuidance) {
                        $aiGroupsApplied = if ($counts.PSObject.Properties["AiSuggestedEditGroupsApplied"]) { [int]$counts.AiSuggestedEditGroupsApplied } else { 0 }
                        $aiGroupsSkipped = if ($counts.PSObject.Properties["AiSuggestedEditGroupsSkipped"]) { [int]$counts.AiSuggestedEditGroupsSkipped } else { 0 }
                        $aiRowsApplied = if ($counts.PSObject.Properties["AiSuggestedEditRowsApplied"]) { [int]$counts.AiSuggestedEditRowsApplied } else { 0 }
                        $aiRowsSkipped = if ($counts.PSObject.Properties["AiSuggestedEditRowsSkipped"]) { [int]$counts.AiSuggestedEditRowsSkipped } else { 0 }
                        Add-Log "AI guidance groups applied: $aiGroupsApplied; skipped/fallback: $aiGroupsSkipped"
                        Add-Log "AI guidance rows applied: $aiRowsApplied; fallback rows: $aiRowsSkipped"
                        $aiManualGroups = if ($counts.PSObject.Properties["AiManualGroupsUsed"]) { [int]$counts.AiManualGroupsUsed } else { 0 }
                        $aiManualRows = if ($counts.PSObject.Properties["AiManualRowsUsed"]) { [int]$counts.AiManualRowsUsed } else { 0 }
                        $aiNoManualRetryGroups = if ($counts.PSObject.Properties["AiNoManualRetryGroups"]) { [int]$counts.AiNoManualRetryGroups } else { 0 }
                        $aiNoManualRetryRows = if ($counts.PSObject.Properties["AiNoManualRetryRows"]) { [int]$counts.AiNoManualRetryRows } else { 0 }
                        $aiNoManualSnippetGroups = if ($counts.PSObject.Properties["AiNoManualSnippetGroups"]) { [int]$counts.AiNoManualSnippetGroups } else { 0 }
                        $aiNoManualSnippetRows = if ($counts.PSObject.Properties["AiNoManualSnippetRows"]) { [int]$counts.AiNoManualSnippetRows } else { 0 }
                        Add-Log "AI used the uploaded manual for $aiManualGroups finding group(s), covering $aiManualRows row(s)."
                        Add-Log "AI handled $aiNoManualRetryGroups finding group(s), covering $aiNoManualRetryRows row(s), without the manual after Azure rejected the manual-context prompt."
                        Add-Log "No matching uploaded-manual guidance was found for $aiNoManualSnippetGroups finding group(s), covering $aiNoManualSnippetRows row(s)."
                    }
                    Add-Log "Elapsed time: $elapsedText"
                    Add-Log "Performance timing summary saved in the run log; top slow steps are also listed on the Run Data tab."
                    if (-not [string]::IsNullOrWhiteSpace([string]$script:GuideRunLogPath)) {
                        Add-Log "Run log saved: $script:GuideRunLogPath"
                    }

                    $aiText = if ([bool]$result.UseAiProjectGuidance) { "On" } else { "Off" }
                    if ([bool]$result.UseAiProjectGuidance) {
                        $aiGroupsApplied = if ($counts.PSObject.Properties["AiSuggestedEditGroupsApplied"]) { [int]$counts.AiSuggestedEditGroupsApplied } else { 0 }
                        $aiGroupsSkipped = if ($counts.PSObject.Properties["AiSuggestedEditGroupsSkipped"]) { [int]$counts.AiSuggestedEditGroupsSkipped } else { 0 }
                        $aiText = "On ($aiGroupsApplied groups applied, $aiGroupsSkipped fallback)"
                    }
                    $message = "Run complete.`r`n`r`nExcel workbook:`r`n$($result.ExportPath)`r`n`r`nRows exported: $($counts.ExportedRows)`r`nElapsed time: $elapsedText`r`nAI guidance: $aiText"
                    if (-not [string]::IsNullOrWhiteSpace([string]$script:GuideRunLogPath)) {
                        $message += "`r`nRun log:`r`n$script:GuideRunLogPath"
                    }
                    $message += "`r`n`r`nOpen the workbook now?"
                    if (-not $script:SuppressGuideRunCompletionDialogs) {
                        $openResult = [System.Windows.Forms.MessageBox]::Show($message, "Run complete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                        if ($openResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                            try { Start-Process -FilePath $result.ExportPath } catch { Add-Log "Could not open workbook: $($_.Exception.Message)" }
                        }
                    }
                }
            }
        }
        finally {
            try { $state.PowerShell.Dispose() } catch { }
            try { $state.Runspace.Close(); $state.Runspace.Dispose() } catch { }
            try { $state.Timer.Dispose() } catch { }
            foreach ($pathToDelete in @($state.ProgressPath, "$($state.ProgressPath).tmp", $state.ResultPath, $state.CancelPath)) {
                if (-not [string]::IsNullOrWhiteSpace($pathToDelete) -and (Test-Path -LiteralPath $pathToDelete)) {
                    try { [System.IO.File]::Delete($pathToDelete) } catch { }
                }
            }
            $script:GuideRunWorker = $null
            $script:GuideRunSettings = $null
            $script:GuideRunStartedAt = $null
            $script:GuideRunCancellationRequested = $false
            $script:GuideRunCancelPath = ""
            $script:GuideRunProgressPath = ""
            $script:GuideRunLogPath = ""
            $script:LastGuideRunLoggedProgress = ""

            if ($script:CloseAfterGuideRunCancel) {
                $script:CloseAfterGuideRunCancel = $false
                try {
                    if ($null -ne $form) { $form.BeginInvoke([System.Action]{ $form.Close() }) | Out-Null }
                }
                catch {
                }
            }
        }
    })

    Set-GuideRunControlsState -Running $true
    if (-not [string]::IsNullOrWhiteSpace([string]$Settings.RunLogPath)) {
        Add-Log "Run log: $($Settings.RunLogPath)"
    }
    Add-Log "Running checks on a temporary copy. Original database will not be changed."
    Add-Log "Performance logging is on. The run log will include step timings, slow SQL previews, and temp-disk context for troubleshooting and speed tuning."
    Add-Log (Get-CleaningPeriodScopeDescription)
    $woodlandRunText = [string]$Settings.WoodlandSpeciesCodesText
    if ([string]::IsNullOrWhiteSpace($woodlandRunText)) {
        Add-Log "Woodland species to exclude from shrinking diameters: none entered."
    }
    else {
        Add-Log "Woodland species to exclude from shrinking diameters: $woodlandRunText"
    }
    Add-Log "Configured range limits: Max IDBH $($Settings.DbhMax); IDBH jump $($Settings.DbhGrowthMax); Max tree height $($Settings.HeightMax); height jump $($Settings.HeightGrowthMax); Max regen stems $($Settings.StemCountMax)."
    Add-Log "Optional tree StemCount check: $(if ($Settings.TreeStemCountChecksEnabled) { 'On' } else { 'Off' })"
    Add-Log "Require DBH for new mortality TreeHistory 2/3: $(if ($Settings.NewMortalityIdbhChecksEnabled) { 'On' } else { 'Off' })"
    Add-Log "Require DBH for old mortality TreeHistory 7: $(if ($Settings.OldMortalityIdbhChecksEnabled) { 'On' } else { 'Off' })"
    Add-Log "Require height for new mortality TreeHistory 2/3: $(if ($Settings.NewMortalityHeightChecksEnabled) { 'On' } else { 'Off' })"
    Add-Log "Require height for old mortality TreeHistory 7: $(if ($Settings.OldMortalityHeightChecksEnabled) { 'On' } else { 'Off' })"
    Add-Log "Total Height Protocol: $(Get-TotalHeightProtocolDisplayText)"
    Add-Log "Total Height legacy code support: Problem 72/74/75 are treated as lean/broken-top/dead-top aliases when the matching 123/127/128 option is on; TreeClass 1/2/3/4/9 and TreeStatus 1/2/3 are treated as live status for TotalHeight checks."
    Add-Log "Total Height subsample minimum count: $(Get-HeightSubsampleMinimumCount); minimum eligible IDBH: $(Get-HeightSubsampleMinimumIdbh); all heights at/above IDBH: $(if (Test-HeightSubsampleAllAtOrAboveEnabled) { Get-HeightSubsampleAllAtOrAboveIdbh } else { 'Off' })"
    Add-Log "Rare species requiring 100% live-tree height: $(Get-HeightRareSpeciesDisplayText)"
    if ([string]::IsNullOrWhiteSpace([string]$Settings.HeightRequiredMinorPlotsText)) {
        Add-Log "Minor plots where tree height is required: none entered; all minor plots are included when height is required."
    }
    else {
        Add-Log "Minor plots where tree height is required: $($Settings.HeightRequiredMinorPlotsText)"
    }
    Add-Log "Regen timber seedling minor plots: $(if ([string]::IsNullOrWhiteSpace([string]$Settings.RegenTimberSeedlingMinorPlotsText)) { 'not enforced' } else { $Settings.RegenTimberSeedlingMinorPlotsText })"
    Add-Log "Regen timber sapling IDBH 20 minor plots: $(if ([string]::IsNullOrWhiteSpace([string]$Settings.RegenTimberSapling20MinorPlotsText)) { 'not enforced' } else { $Settings.RegenTimberSapling20MinorPlotsText })"
    Add-Log "Regen timber sapling IDBH 40 minor plots: $(if ([string]::IsNullOrWhiteSpace([string]$Settings.RegenTimberSapling40MinorPlotsText)) { 'not enforced' } else { $Settings.RegenTimberSapling40MinorPlotsText })"
    Add-Log "Regen woodland seedling minor plots: $(if ([string]::IsNullOrWhiteSpace([string]$Settings.RegenWoodlandSeedlingMinorPlotsText)) { 'not enforced' } else { $Settings.RegenWoodlandSeedlingMinorPlotsText })"
    Add-Log "Regen woodland sapling IDBH 20 minor plots: $(if ([string]::IsNullOrWhiteSpace([string]$Settings.RegenWoodlandSapling20MinorPlotsText)) { 'not enforced' } else { $Settings.RegenWoodlandSapling20MinorPlotsText })"
    Add-Log "Regen woodland sapling IDBH 40 minor plots: $(if ([string]::IsNullOrWhiteSpace([string]$Settings.RegenWoodlandSapling40MinorPlotsText)) { 'not enforced' } else { $Settings.RegenWoodlandSapling40MinorPlotsText })"
    if (([string]::IsNullOrWhiteSpace($woodlandRunText)) -and
        ((-not [string]::IsNullOrWhiteSpace([string]$Settings.RegenWoodlandSeedlingMinorPlotsText)) -or
         (-not [string]::IsNullOrWhiteSpace([string]$Settings.RegenWoodlandSapling20MinorPlotsText)) -or
         (-not [string]::IsNullOrWhiteSpace([string]$Settings.RegenWoodlandSapling40MinorPlotsText)))) {
        Add-Log "Woodland regen minor-plot rules need woodland species codes above; woodland regen rules will not be enforced without them."
    }
    if ($Settings.UseAiProjectGuidance) {
        Add-Log "AI guidance is on. The export will ask for AIMessage guidance using coded rules, template metadata, and manual excerpts when loaded."
        Add-Log "Compact AI export prompts: $(if ($Settings.AiCompactExport) { 'On' } else { 'Off' })"
        if (-not [string]::IsNullOrWhiteSpace([string]$Settings.ProjectManualText)) {
            $manualName = [System.IO.Path]::GetFileName([string]$Settings.ProjectManualPath)
            if ([string]::IsNullOrWhiteSpace($manualName)) { $manualName = "loaded manual" }
            Add-Log "Project manual loaded for AI export: $manualName"
        }
        else {
            Add-Log "No uploaded project manual loaded for AI export."
        }
    }
    Set-AppProgress -Status "Starting run..." -Percent 1
    $timer.Start()
}

function Run-CfiWorkbookChecks {
    if ([string]::IsNullOrWhiteSpace($script:DbPath) -or [string]::IsNullOrWhiteSpace($script:ConnectionString)) {
        [System.Windows.Forms.MessageBox]::Show("Connect to a database first.", $script:AppName) | Out-Null
        return
    }

    $forceRunWithoutAi = Confirm-RunWithoutAiIfNeeded
    if ($null -eq $forceRunWithoutAi) { return }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Excel workbook (*.xlsx)|*.xlsx"
    $dialog.Title = "Save run workbook"
    $dialog.FileName = Get-DefaultRunWorkbookFileName
    if (-not [string]::IsNullOrWhiteSpace($script:DbPath)) {
        $dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:DbPath)
    }

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        $dialog.Dispose()
        return
    }

    $exportPath = $dialog.FileName
    $dialog.Dispose()

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will run workbook-style cleaning checks on a temporary copy of the database. Your original Access database will not be changed. The results will be saved to Excel. Continue?",
        "Run",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        $settings = Get-GuideRunSettingsSnapshot -ExportPath $exportPath -ForceDisableAiProjectGuidance:([bool]$forceRunWithoutAi)
        Start-GuideChecksBackgroundRun -Settings $settings
    }
    catch {
        Set-AppProgress -Status "Run failed." -Percent 0
        $failureSettings = [pscustomobject]@{
            SourcePath = [string]$script:DbPath
            ExportPath = [string]$exportPath
            RunLogPath = [string]$runLogPath
            StartedAt = (Get-Date).ToString("o")
            UseAiProjectGuidance = [bool](Test-AiProjectGuidanceSelected)
            AiModel = [string]$script:AiDefaultModel
            ProjectManualPath = [string]$script:ProjectManualPath
        }
        Initialize-GuideRunLog -Settings $failureSettings
        Write-GuideRunFailureLog -Status "Failed during setup" -Message $_.Exception.Message -Detail ($_ | Out-String) -Settings $failureSettings
        Add-Log "Run failed: $($_.Exception.Message)"
        if (-not [string]::IsNullOrWhiteSpace([string]$runLogPath)) {
            Add-Log "Failure details saved to run log: $runLogPath"
        }
        $failureMessage = $_.Exception.Message
        if (-not [string]::IsNullOrWhiteSpace([string]$runLogPath)) {
            $failureMessage += "`r`n`r`nRun log:`r`n$runLogPath"
        }
        [System.Windows.Forms.MessageBox]::Show($failureMessage, "Run failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $script:GuideRunLogPath = ""
    }
}

function Refresh-Tables {
    $connection = Open-AccessConnection
    try {
        $script:PeriodScopeConditionCache = @{}
        $tableCombo.Items.Clear()
        $tables = Get-UserTables $connection
        foreach ($name in $tables) {
            [void]$tableCombo.Items.Add($name)
        }

        if ($tableCombo.Items.Count -gt 0) {
            $tableCombo.SelectedIndex = 0
        }

        Add-Log "Loaded $($tableCombo.Items.Count) user tables."
        Set-PeriodScopeDefaultsFromConnection -Connection $connection
        Add-Log (Get-CleaningPeriodScopeDescription)
    }
    finally {
        $connection.Close()
        $connection.Dispose()
    }
}

function Refresh-ColumnsForSelectedTable {
    $tableName = [string]$tableCombo.SelectedItem
    if ([string]::IsNullOrWhiteSpace($tableName)) { return }

    $connection = Open-AccessConnection
    try {
        $columns = @(Get-TableColumns -Connection $connection -TableName $tableName)
        $script:ColumnsByTable[$tableName] = $columns

        $species = Find-CandidateColumn -Columns $columns -Patterns @("^(species|spp|sp|spcd|speciescode)$", "species", "spcode") -TextOnly
        $dbh = Find-CandidateColumn -Columns $columns -Patterns @("^idbh$") -NumericOnly
        $height = Find-CandidateColumn -Columns $columns -Patterns @("^totalheight$", "^height$", "^ht$", "^tht$", "treeheight", "totalheight") -NumericOnly
        $plot = Find-CandidateColumn -Columns $columns -Patterns @("^plotmeaskey$", "^plotkey$", "^plotid$", "^plotnumber$", "^plot$", "plotid", "plotnumber", "stand", "unit")
        $tree = Find-CandidateColumn -Columns $columns -Patterns @("^treemeaskey$", "^regenmeaskey$", "^treekey$", "^treeid$", "^treenumber$", "^tree$", "tag", "stem")

        Set-ComboItems -Combo $speciesCombo -Columns $columns -SelectedName $species
        Set-ComboItems -Combo $dbhCombo -Columns $columns -SelectedName $dbh
        Set-ComboItems -Combo $heightCombo -Columns $columns -SelectedName $height
        Set-ComboItems -Combo $plotCombo -Columns $columns -SelectedName $plot
        Set-ComboItems -Combo $treeCombo -Columns $columns -SelectedName $tree

        Add-Log "Loaded $($columns.Count) columns for table '$tableName'."
    }
    finally {
        $connection.Close()
        $connection.Dispose()
    }
}

function Show-Preview {
    $tableName = [string]$tableCombo.SelectedItem
    if ([string]::IsNullOrWhiteSpace($tableName)) {
        [System.Windows.Forms.MessageBox]::Show("Choose a table first.", $script:AppName) | Out-Null
        return
    }

    $connection = Open-AccessConnection
    try {
        $grid.DataSource = Get-DataTable -Connection $connection -Sql "SELECT TOP 100 * FROM $(Quote-Name $tableName)"
        $rowCount = Get-Scalar -Connection $connection -Sql "SELECT Count(*) FROM $(Quote-Name $tableName)"
        Add-Log "Previewing first 100 rows from '$tableName'. Total rows: $rowCount."
    }
    finally {
        $connection.Close()
        $connection.Dispose()
    }
}

function Clean-SelectedTable {
    $tableName = [string]$tableCombo.SelectedItem
    if ([string]::IsNullOrWhiteSpace($tableName)) {
        [System.Windows.Forms.MessageBox]::Show("Choose a table first.", $script:AppName) | Out-Null
        return
    }

    $columns = @($script:ColumnsByTable[$tableName])
    if ($columns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Load the table columns first.", $script:AppName) | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will create a backup, clean selected text values, add or update a NeedsReview field, and append findings to InventoryCleanAudit. No rows will be deleted. Continue?",
        "Clean selected table",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $connection = $null
    $transaction = $null

    try {
        $backupPath = Backup-Database -Path $script:DbPath
        Add-Log "Backup created: $backupPath"

        $connection = Open-AccessConnection
        Ensure-AuditTable -Connection $connection
        Ensure-NeedsReviewColumn -Connection $connection -TableName $tableName

        $transaction = $connection.BeginTransaction()

        $trimCount = 0
        $nullCount = 0
        $speciesCount = 0
        $codeAuditCount = 0

        if ($trimTextCheck.Checked -or $blankNullCheck.Checked) {
            foreach ($column in $columns) {
                if (-not (Test-TextColumn $column.Type)) { continue }
                $field = Quote-Name $column.Name

                if ($trimTextCheck.Checked) {
                    $sql = "UPDATE $(Quote-Name $tableName) SET $field = Trim($field) WHERE $field Is Not Null AND $field <> Trim($field)"
                    $affected = Invoke-NonQuery -Connection $connection -Transaction $transaction -Sql $sql
                    if ($affected -gt 0) { $trimCount += $affected }
                }

                if ($blankNullCheck.Checked) {
                    $sql = "UPDATE $(Quote-Name $tableName) SET $field = Null WHERE $field Is Not Null AND Trim($field) = ''"
                    $affected = Invoke-NonQuery -Connection $connection -Transaction $transaction -Sql $sql
                    if ($affected -gt 0) { $nullCount += $affected }
                }
            }
        }

        $speciesField = Get-SelectedField $speciesCombo
        if ($speciesUpperCheck.Checked -and $speciesField) {
            $speciesColumn = Get-ColumnByName -Columns $columns -Name $speciesField
            if ($speciesColumn -and (Test-TextColumn $speciesColumn.Type)) {
                $field = Quote-Name $speciesField
                $sql = "UPDATE $(Quote-Name $tableName) SET $field = UCase(Trim($field)) WHERE $field Is Not Null AND $field <> UCase(Trim($field))"
                $affected = Invoke-NonQuery -Connection $connection -Transaction $transaction -Sql $sql
                if ($affected -gt 0) { $speciesCount += $affected }
            }
            else {
                Add-Log "Skipped species uppercase because '$speciesField' is not a text field."
            }
        }

        $codeAuditCount = Add-CodeValidationAudit `
            -Connection $connection `
            -Transaction $transaction `
            -TableName $tableName `
            -Columns $columns

        $plotField = Get-SelectedField $plotCombo
        $treeField = Get-SelectedField $treeCombo
        $recordLabel = Get-RecordLabelExpression -PlotField $plotField -TreeField $treeField

        $dbhAuditCount = 0
        $heightAuditCount = 0
        if ($rangeCheck.Checked) {
            $dbhField = Get-SelectedField $dbhCombo
            if ($dbhField -and (Test-AppColumnFieldActiveForReview -Connection $connection -TableName $tableName -FieldName $dbhField)) {
                $dbhRangeMessage = Add-FieldManualTipToMessage `
                    -Message "Entered IDBH values must be greater than zero and no more than the configured maximum." `
                    -TableName $tableName `
                    -FieldName $dbhField
                if ($tableName -eq "TreeMeasurements" -and (Test-ColumnExists -Columns $columns -Name "TreeHistory")) {
                    $dbhFieldQuoted = Quote-Name $dbhField
                    $dbhMaxSql = ([decimal]$dbhMax.Value).ToString([System.Globalization.CultureInfo]::InvariantCulture)
                    $missingRequiredCondition = Get-TreeIdbhMissingRequiredCondition -TreeHistoryField "[TreeHistory]" -IdbhField $dbhFieldQuoted
                    $missingDbhMessage = Add-FieldManualTipToMessage `
                        -Message "TreeMeasurements.IDBH is blank, but this TreeHistory status is selected in Run options as requiring DBH." `
                        -TableName $tableName `
                        -FieldName $dbhField
                    $dbhAuditCount += Add-ConditionAudit `
                        -Connection $connection `
                        -Transaction $transaction `
                        -TableName $tableName `
                        -RuleName "Missing DBH" `
                        -FieldName $dbhField `
                        -ObservedExpression $dbhFieldQuoted `
                        -Message $missingDbhMessage `
                        -RecordLabelExpression $recordLabel `
                        -Condition $missingRequiredCondition
                    $dbhAuditCount += Add-ConditionAudit `
                        -Connection $connection `
                        -Transaction $transaction `
                        -TableName $tableName `
                        -RuleName "IDBH range check" `
                        -FieldName $dbhField `
                        -ObservedExpression $dbhFieldQuoted `
                        -Message $dbhRangeMessage `
                        -RecordLabelExpression $recordLabel `
                        -Condition "$dbhFieldQuoted Is Not Null AND ($dbhFieldQuoted <= 0 OR $dbhFieldQuoted > $dbhMaxSql)"
                }
                else {
                    $dbhAuditCount = Add-RangeAudit `
                        -Connection $connection `
                        -Transaction $transaction `
                        -TableName $tableName `
                        -FieldName $dbhField `
                        -MaximumValue $dbhMax.Value `
                        -RuleName "IDBH range check" `
                        -Message $dbhRangeMessage `
                        -RecordLabelExpression $recordLabel
                }
            }

            $heightField = Get-SelectedField $heightCombo
            if ($heightField -and (Test-AppColumnFieldActiveForReview -Connection $connection -TableName $tableName -FieldName $heightField)) {
                $heightMessage = Add-FieldManualTipToMessage `
                    -Message "Height is missing, zero, negative, or above the configured maximum." `
                    -TableName $tableName `
                    -FieldName $heightField
                $heightAuditCount = Add-RangeAudit `
                    -Connection $connection `
                    -Transaction $transaction `
                    -TableName $tableName `
                    -FieldName $heightField `
                    -MaximumValue $heightMax.Value `
                    -RuleName "Height range check" `
                    -Message $heightMessage `
                    -RecordLabelExpression $recordLabel
            }
        }

        $duplicateCount = 0
        if ($duplicateCheck.Checked) {
            $duplicateCount = Add-DuplicateAudit `
                -Connection $connection `
                -Transaction $transaction `
                -TableName $tableName `
                -PlotField $plotField `
                -TreeField $treeField
        }

        $transaction.Commit()
        $transaction = $null

        Add-Log "Clean complete for '$tableName'."
        Add-Log "Trimmed text updates: $trimCount"
        Add-Log "Blank text converted to Null: $nullCount"
        Add-Log "Species code updates: $speciesCount"
        Add-Log "Invalid CFI code audit records: $codeAuditCount"
        Add-Log "IDBH audit records: $dbhAuditCount"
        Add-Log "Height audit records: $heightAuditCount"
        Add-Log "Duplicate key audit records: $duplicateCount"

        Show-Preview
    }
    catch {
        if ($null -ne $transaction) {
            try { $transaction.Rollback() } catch { }
        }
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Cleaning failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        Add-Log "Cleaning failed: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $connection) {
            $connection.Close()
            $connection.Dispose()
        }
    }
}

function Get-TableSampleTextForAi {
    param(
        [System.Data.OleDb.OleDbConnection]$Connection,
        [string]$TableName,
        [int]$MaxRows = 5,
        [int]$MaxColumns = 12
    )

    try {
        $table = Get-DataTable -Connection $Connection -Sql "SELECT TOP $MaxRows * FROM $(Quote-Name $TableName)"
        if ($table.Rows.Count -eq 0) { return "No sample rows." }

        $columns = @($table.Columns | Select-Object -First $MaxColumns)
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($row in $table.Rows) {
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($column in $columns) {
                $value = Get-DataRowText -Row $row -ColumnName $column.ColumnName
                if ($value.Length -gt 80) { $value = $value.Substring(0, 77) + "..." }
                [void]$parts.Add(("{0}={1}" -f $column.ColumnName, $value))
            }
            [void]$lines.Add([string]::Join("; ", [string[]]$parts.ToArray()))
        }
        return [string]::Join("`r`n", [string[]]$lines.ToArray())
    }
    catch {
        return "Could not read sample rows: $($_.Exception.Message)"
    }
}

function Get-ConnectedDatabaseOverviewForAi {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine("Connected database overview:")
    [void]$builder.AppendLine("Path: $script:DbPath")

    $targetTables = @(Get-CfiWorkbookTargetTables -Connection $Connection)
    if ($targetTables.Count -eq 0) {
        $targetTables = @(Get-UserTables -Connection $Connection | Select-Object -First 20)
    }

    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("CFI/core table counts:")
    foreach ($tableName in $targetTables) {
        try {
            $count = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM $(Quote-Name $tableName)"
            [void]$builder.AppendLine("- ${tableName}: $count rows")
        }
        catch {
            [void]$builder.AppendLine("- ${tableName}: could not count rows")
        }
    }

    $selectedTable = $null
    try {
        if ($null -ne $tableCombo -and $tableCombo.SelectedItem) {
            $selectedTable = [string]$tableCombo.SelectedItem
        }
    }
    catch {
        $selectedTable = $null
    }

    if ([string]::IsNullOrWhiteSpace($selectedTable)) {
        foreach ($candidate in @("TreeMeasurements", "Trees", "RegenMeasurements", "PlotMeasurements")) {
            if ($targetTables -contains $candidate) {
                $selectedTable = $candidate
                break
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($selectedTable)) {
        [void]$builder.AppendLine("")
        [void]$builder.AppendLine("Selected/focus table: $selectedTable")
        try {
            $columns = @(Get-TableColumns -Connection $Connection -TableName $selectedTable)
            $columnText = [string]::Join(", ", [string[]]@($columns | ForEach-Object { $_.Name } | Select-Object -First 40))
            [void]$builder.AppendLine("Columns: $columnText")
        }
        catch {
            [void]$builder.AppendLine("Columns: could not read columns")
        }
        [void]$builder.AppendLine("Sample rows:")
        [void]$builder.AppendLine((Get-TableSampleTextForAi -Connection $Connection -TableName $selectedTable -MaxRows 5 -MaxColumns 14))
    }

    return $builder.ToString()
}

function Get-GuideCheckContextForAi {
    param(
        [string]$SourcePath,
        [string]$Password
    )

    if ([string]::IsNullOrWhiteSpace($SourcePath) -or -not (Test-Path -LiteralPath $SourcePath)) {
        return "Run context: no connected database file was available."
    }

    $tempDirectory = Get-AppTempDirectory
    [System.IO.Directory]::CreateDirectory($tempDirectory) | Out-Null
    $extension = [System.IO.Path]::GetExtension($SourcePath)
    $tempPath = Join-Path $tempDirectory ("ai_context_" + [guid]::NewGuid().ToString("N") + $extension)
    $lockPath = [System.IO.Path]::ChangeExtension($tempPath, ".ldb")
    if ($extension -and $extension.ToLowerInvariant() -eq ".accdb") {
        $lockPath = [System.IO.Path]::ChangeExtension($tempPath, ".laccdb")
    }

    $connection = $null
    try {
        Copy-Item -LiteralPath $SourcePath -Destination $tempPath -Force
        $tempConnectionString = New-AccessConnectionString -Path $tempPath -Password $Password
        $connection = New-Object System.Data.OleDb.OleDbConnection($tempConnectionString)
        $connection.Open()

        $counts = Invoke-CfiGuideChecks -Connection $connection
        $builder = New-Object System.Text.StringBuilder
        [void]$builder.AppendLine("Fresh run findings from a temporary copy:")
        [void]$builder.AppendLine(("- Workbook counts: {0}" -f $counts.WorkbookCounts))
        [void]$builder.AppendLine(("- R-script reference: {0}" -f $counts.RScriptReference))
        [void]$builder.AppendLine(("- Invalid codes: {0}" -f $counts.InvalidCodes))
        [void]$builder.AppendLine(("- Range/tree species/crown/radial/required plot checks: {0}" -f $counts.Ranges))
        if ($counts.PSObject.Properties["PlotStatusProgression"]) {
            [void]$builder.AppendLine(("- PlotStatus progression: {0}" -f $counts.PlotStatusProgression))
        }
        if ($counts.PSObject.Properties["InactivePlotData"]) {
            [void]$builder.AppendLine(("- Plot data on not-measured/dropped plots: {0}" -f $counts.InactivePlotData))
        }
        if ($counts.PSObject.Properties["InactivePlotTreeData"]) {
            [void]$builder.AppendLine(("- Tree data on not-measured/dropped plots: {0}" -f $counts.InactivePlotTreeData))
        }
        if ($counts.PSObject.Properties["InactivePlotRegenData"]) {
            [void]$builder.AppendLine(("- Regen data on not-measured/dropped plots: {0}" -f $counts.InactivePlotRegenData))
        }
        [void]$builder.AppendLine(("- Remeasurement checks: {0}" -f $counts.Remeasurement))
        if ($counts.PSObject.Properties["WoodlandShrinkSkipVerification"]) {
            [void]$builder.AppendLine(("- Woodland DBH shrink exclusion: {0}" -f $counts.WoodlandShrinkSkipVerification.Message))
        }

        $tables = @(Get-UserTables -Connection $connection -IncludeAudit)
        if ($tables -contains "InventoryCleanAudit") {
            $summary = Get-DataTable -Connection $connection -Sql "SELECT [FieldName], [RuleName], Count(*) AS [FindingCount] FROM [InventoryCleanAudit] GROUP BY [FieldName], [RuleName] ORDER BY Count(*) DESC, [FieldName], [RuleName]"
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine("Top finding groups:")
            $groupLimit = [Math]::Min(25, $summary.Rows.Count)
            for ($i = 0; $i -lt $groupLimit; $i++) {
                $row = $summary.Rows[$i]
                [void]$builder.AppendLine(("- Field={0}; Rule={1}; Count={2}" -f $row["FieldName"], $row["RuleName"], $row["FindingCount"]))
            }

            $recent = Get-DataTable -Connection $connection -Sql "SELECT TOP 40 [TableName], [RuleName], [RecordLabel], [FieldName], [ObservedValue], [Message] FROM [InventoryCleanAudit] ORDER BY [AuditId]"
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine("Example findings:")
            foreach ($row in $recent.Rows) {
                $message = Get-DataRowText -Row $row -ColumnName "Message"
                if ($message.Length -gt 240) { $message = $message.Substring(0, 237) + "..." }
                [void]$builder.AppendLine(("- Table={0}; Rule={1}; Field={2}; Record={3}; Observed={4}; Message={5}" -f $row["TableName"], $row["RuleName"], $row["FieldName"], $row["RecordLabel"], $row["ObservedValue"], $message))
            }
        }

        return $builder.ToString()
    }
    catch {
        return "Run context failed: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $connection) {
            $connection.Close()
            $connection.Dispose()
        }
        foreach ($pathToDelete in @($lockPath, $tempPath)) {
            if (-not [string]::IsNullOrWhiteSpace($pathToDelete) -and (Test-Path -LiteralPath $pathToDelete)) {
                try { [System.IO.File]::Delete($pathToDelete) } catch { }
            }
        }
    }
}

function Get-ExistingAuditContextForAi {
    param([System.Data.OleDb.OleDbConnection]$Connection)

    try {
        $tables = @(Get-UserTables -Connection $Connection -IncludeAudit)
        if (-not ($tables -contains "InventoryCleanAudit")) {
            return "Run findings: none available in the connected database."
        }

        $count = Get-CountValue -Connection $Connection -Sql "SELECT Count(*) FROM [InventoryCleanAudit]"
        if ($count -eq 0) {
            return "Run findings: InventoryCleanAudit exists, but it has no rows."
        }

        $builder = New-Object System.Text.StringBuilder
        [void]$builder.AppendLine("Existing InventoryCleanAudit rows in the connected database: $count")
        $summary = Get-DataTable -Connection $Connection -Sql "SELECT [FieldName], [RuleName], Count(*) AS [FindingCount] FROM [InventoryCleanAudit] GROUP BY [FieldName], [RuleName] ORDER BY Count(*) DESC, [FieldName], [RuleName]"
        $groupLimit = [Math]::Min(20, $summary.Rows.Count)
        for ($i = 0; $i -lt $groupLimit; $i++) {
            $row = $summary.Rows[$i]
            [void]$builder.AppendLine(("- Field={0}; Rule={1}; Count={2}" -f $row["FieldName"], $row["RuleName"], $row["FindingCount"]))
        }
        return $builder.ToString()
    }
    catch {
        return "Existing audit context could not be read: $($_.Exception.Message)"
    }
}

function Get-DatabaseContextForAi {
    param([bool]$IncludeGuideChecks = $false)

    if ([string]::IsNullOrWhiteSpace($script:ConnectionString) -or [string]::IsNullOrWhiteSpace($script:DbPath)) {
        return "No database is connected."
    }

    $connection = $null
    try {
        $connection = Open-AccessConnection
        $builder = New-Object System.Text.StringBuilder
        [void]$builder.AppendLine((Get-ConnectedDatabaseOverviewForAi -Connection $connection))
        [void]$builder.AppendLine("")
        [void]$builder.AppendLine((Get-CleaningRulesContextForAi -Connection $connection))
        [void]$builder.AppendLine("")

        if ($IncludeGuideChecks) {
            $password = ""
            try {
                if ($null -ne $passwordBox) { $password = $passwordBox.Text }
            }
            catch {
                $password = ""
            }
            [void]$builder.AppendLine((Get-GuideCheckContextForAi -SourcePath $script:DbPath -Password $password))
        }
        else {
            [void]$builder.AppendLine((Get-ExistingAuditContextForAi -Connection $connection))
            [void]$builder.AppendLine("Fresh checks were not run for this AI message.")
        }
        [void]$builder.AppendLine("")
        $projectManualContext = Get-ProjectManualContextForAi -MaxLength 10000
        if (-not [string]::IsNullOrWhiteSpace($projectManualContext)) {
            [void]$builder.AppendLine($projectManualContext)
            [void]$builder.AppendLine("")
        }
        if ($IncludeGuideChecks) {
            [void]$builder.AppendLine("Privacy/safety note: the app read the connected database locally and sent this compact context, not the full Access file. Fresh checks were run on a temporary copy so the source database was not changed.")
        }
        else {
            [void]$builder.AppendLine("Privacy/safety note: the app read the connected database locally and sent this compact context, not the full Access file. Fresh checks were not run for this AI message.")
        }
        return $builder.ToString()
    }
    catch {
        return "Could not load database context: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $connection) {
            $connection.Close()
            $connection.Dispose()
        }
    }
}

function Get-OpenAiResponseText {
    param([object]$Response)

    if ($null -eq $Response) { return "" }
    if ($Response -is [string]) {
        $textResponse = ([string]$Response).Trim()
        if ([string]::IsNullOrWhiteSpace($textResponse) -or $textResponse -eq '""') { return "" }
        if ($textResponse.StartsWith('"') -and $textResponse.EndsWith('"')) {
            try {
                $decoded = $textResponse | ConvertFrom-Json
                if ($decoded -is [string]) { return [string]$decoded }
            }
            catch {
            }
        }
        return $textResponse
    }

    if ($Response.PSObject.Properties["output_text"] -and -not [string]::IsNullOrWhiteSpace([string]$Response.output_text)) {
        return [string]$Response.output_text
    }

    if ($Response.PSObject.Properties["choices"] -and $Response.choices.Count -gt 0) {
        $choice = $Response.choices[0]
        if ($choice.PSObject.Properties["message"] -and $choice.message.PSObject.Properties["content"]) {
            $content = $choice.message.content
            if ($content -is [array]) {
                $parts = New-Object System.Collections.Generic.List[string]
                foreach ($item in @($content)) {
                    if ($item.PSObject.Properties["text"]) {
                        [void]$parts.Add([string]$item.text)
                    }
                    elseif ($item.PSObject.Properties["content"]) {
                        [void]$parts.Add([string]$item.content)
                    }
                    else {
                        [void]$parts.Add([string]$item)
                    }
                }
                return [string]::Join("`r`n", [string[]]$parts.ToArray())
            }
            return [string]$content
        }
        if ($choice.PSObject.Properties["text"]) {
            return [string]$choice.text
        }
    }

    if ($Response.PSObject.Properties["output"]) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($item in @($Response.output)) {
            if (-not $item.PSObject.Properties["content"]) { continue }
            foreach ($content in @($item.content)) {
                if ($content.PSObject.Properties["text"]) {
                    $textValue = $content.text
                    if ($textValue.PSObject.Properties["value"]) {
                        [void]$parts.Add([string]$textValue.value)
                    }
                    else {
                        [void]$parts.Add([string]$textValue)
                    }
                }
                elseif ($content.PSObject.Properties["content"]) {
                    [void]$parts.Add([string]$content.content)
                }
            }
        }
        if ($parts.Count -gt 0) { return [string]::Join("`r`n", [string[]]$parts.ToArray()) }
    }

    $fallback = ($Response | ConvertTo-Json -Depth 12)
    if ($fallback -eq '""') { return "" }
    return $fallback
}

function Invoke-AiConnectionTestRequest {
    param(
        [string]$Endpoint,
        [string]$ApiKey,
        [string]$Model
    )

    if ([string]::IsNullOrWhiteSpace($ApiKey)) { throw "Enter an API key." }
    if ([string]::IsNullOrWhiteSpace($Model)) { throw "Enter a model." }
    $target = Get-AiRequestTarget -Endpoint $Endpoint -Model $Model

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $headers = New-AiRequestHeaders -ApiKey $ApiKey -IsAzure ([bool]$target.IsAzure)

    if ($target.UseChatCompletions) {
        $body = New-AiChatBody `
            -Model $Model `
            -SystemText "This is a connection test. Reply with OK." `
            -UserText "Connection test." `
            -IsAzure:([bool]$target.IsAzure) `
            -IncludeModelInBody:([bool]$target.IncludeModelInBody)
    }
    else {
        $body = @{
            model = $Model
            instructions = (ConvertTo-AiSafeText -Value "This is a connection test. Reply with OK.")
            input = (ConvertTo-AiSafeText -Value "Connection test.")
            max_output_tokens = 20
        }
    }

    $json = $body | ConvertTo-Json -Depth 10
    $jsonBody = ConvertTo-Utf8JsonBody -Json $json
    try {
        $response = Invoke-RestMethod -Method Post -Uri ([string]$target.Uri) -Headers $headers -ContentType "application/json; charset=utf-8" -Body $jsonBody -TimeoutSec 45
    }
    catch {
        throw (Get-FriendlyAiRequestError -Exception $_.Exception -ErrorRecord $_ -Endpoint ([string]$target.Uri))
    }

    $text = Get-OpenAiResponseText -Response $response
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "Connection succeeded, but the model returned an empty response."
    }

    return "Connection succeeded. Model response: " + (Limit-TextLength -Text $text.Trim() -MaxLength 200)
}

function Invoke-AiChatRequest {
    param(
        [string]$Endpoint,
        [string]$ApiKey,
        [string]$Model,
        [string]$Question,
        [string]$AuditContext,
        [string]$PriorChat
    )

    if ([string]::IsNullOrWhiteSpace($ApiKey)) { throw "Enter an API key." }
    if ([string]::IsNullOrWhiteSpace($Model)) { throw "Enter a model." }
    $target = Get-AiRequestTarget -Endpoint $Endpoint -Model $Model

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $instructions = @"
You are a careful forest inventory data QA assistant inside a desktop cleaning tool.
Help the user interpret the connected Access database, coded cleaning rules, project template metadata, run findings, table counts, schema, sample records, and cleaning risks.
Provide concise insights, prioritized cleaning recommendations, and Access SQL the user can run to verify issues.
Format every chat answer as Markdown. Use short headings, bullets, and fenced SQL code blocks when giving Access SQL. Keep the tone concise and practical.
Do not invent exact corrected values. When a value needs source verification, say so plainly.
Treat the coded cleaning rules and template metadata in the context as authoritative. Do not tell the user to update the Access database unless the correction is supported by the database context, run findings, or original field record.
Prefer review-first language for biological or logical anomalies such as shrinking IDBH, height changes, Lazarus tree histories, and missing measurement rows. Shrinking IDBH findings are meant for live timber trees; woodland species codes entered in the app are allowed to shrink and should not be treated as errors.
"@
    $inputText = @"
Current connected-database context:
$AuditContext

Recent chat:
$PriorChat

User question:
$Question
"@

    $headers = New-AiRequestHeaders -ApiKey $ApiKey -IsAzure ([bool]$target.IsAzure)

    if ($target.UseChatCompletions) {
        $body = New-AiChatBody `
            -Model $Model `
            -SystemText $instructions `
            -UserText $inputText `
            -IsAzure:([bool]$target.IsAzure) `
            -IncludeModelInBody:([bool]$target.IncludeModelInBody)
    }
    else {
        $body = @{
            model = $Model
            instructions = (ConvertTo-AiSafeText -Value $instructions)
            input = (ConvertTo-AiSafeText -Value $inputText)
        }
    }

    $json = $body | ConvertTo-Json -Depth 10
    $jsonBody = ConvertTo-Utf8JsonBody -Json $json
    try {
        $response = Invoke-RestMethod -Method Post -Uri ([string]$target.Uri) -Headers $headers -ContentType "application/json; charset=utf-8" -Body $jsonBody -TimeoutSec 90
    }
    catch {
        throw (Get-FriendlyAiRequestError -Exception $_.Exception -ErrorRecord $_ -Endpoint ([string]$target.Uri))
    }
    $text = Get-OpenAiResponseText -Response $response
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "The model returned an empty response. Use Test connection to confirm the endpoint, model/deployment name, and API key."
    }
    return $text
}

function Show-AiHelperDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "AI guidance"
    $dialog.StartPosition = "CenterParent"
    $dialog.Size = New-Object System.Drawing.Size(900, 720)
    $dialog.MinimumSize = New-Object System.Drawing.Size(780, 620)
    $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dialog.BackColor = $script:DadaColorBackground
    try {
        if ($null -ne $script:AppIcon) { $dialog.Icon = $script:AppIcon }
    }
    catch {
    }

    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = "Top"
    $headerPanel.Height = 84
    $headerPanel.BackColor = $script:DadaColorSurface
    $dialog.Controls.Add($headerPanel)

    $helperIcon = New-Object System.Windows.Forms.PictureBox
    $helperIcon.Location = New-Object System.Drawing.Point(18, 17)
    $helperIcon.Size = New-Object System.Drawing.Size(48, 48)
    $helperIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $helperIcon.BackColor = $script:DadaColorSurface
    try {
        if ($null -ne $script:AppIcon) { $helperIcon.Image = $script:AppIcon.ToBitmap() }
    }
    catch {
    }
    $headerPanel.Controls.Add($helperIcon)

    $helperTitleLabel = New-Object System.Windows.Forms.Label
    $helperTitleLabel.Text = "AI guidance"
    $helperTitleLabel.Location = New-Object System.Drawing.Point(78, 16)
    $helperTitleLabel.Size = New-Object System.Drawing.Size(420, 28)
    $helperTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14)
    $helperTitleLabel.ForeColor = $script:DadaColorAccent
    $headerPanel.Controls.Add($helperTitleLabel)

    $helperSubtitleLabel = New-Object System.Windows.Forms.Label
    $helperSubtitleLabel.Text = "Optional project-specific help for database questions and Excel export guidance."
    $helperSubtitleLabel.Location = New-Object System.Drawing.Point(80, 46)
    $helperSubtitleLabel.Size = New-Object System.Drawing.Size(760, 22)
    $helperSubtitleLabel.Anchor = "Top,Left,Right"
    $helperSubtitleLabel.ForeColor = $script:DadaColorMutedText
    $headerPanel.Controls.Add($helperSubtitleLabel)

    $settingsGroup = New-Object System.Windows.Forms.GroupBox
    $settingsGroup.Text = "Connection"
    $settingsGroup.Location = New-Object System.Drawing.Point(14, 96)
    $settingsGroup.Size = New-Object System.Drawing.Size(854, 148)
    $settingsGroup.Anchor = "Top,Left,Right"
    $settingsGroup.BackColor = $script:DadaColorSurface
    $dialog.Controls.Add($settingsGroup)

    $enabledCheck = New-Object System.Windows.Forms.CheckBox
    $enabledCheck.Text = "AI helper on"
    $enabledCheck.Location = New-Object System.Drawing.Point(16, 28)
    $enabledCheck.Size = New-Object System.Drawing.Size(130, 24)
    $enabledCheck.Checked = $aiEnabledCheck.Checked
    $settingsGroup.Controls.Add($enabledCheck)

    $includeGuideChecksCheck = New-Object System.Windows.Forms.CheckBox
    $includeGuideChecksCheck.Text = "Include fresh run findings"
    $includeGuideChecksCheck.Location = New-Object System.Drawing.Point(158, 28)
    $includeGuideChecksCheck.Size = New-Object System.Drawing.Size(220, 24)
    $includeGuideChecksCheck.Checked = $false
    $settingsGroup.Controls.Add($includeGuideChecksCheck)

    $testConnectionButton = New-Object System.Windows.Forms.Button
    $testConnectionButton.Text = "Test connection"
    $testConnectionButton.Location = New-Object System.Drawing.Point(700, 24)
    $testConnectionButton.Size = New-Object System.Drawing.Size(130, 30)
    $testConnectionButton.Anchor = "Top,Right"
    $settingsGroup.Controls.Add($testConnectionButton)

    $endpointLabel = New-Object System.Windows.Forms.Label
    $endpointLabel.Text = "Endpoint"
    $endpointLabel.Location = New-Object System.Drawing.Point(16, 64)
    $endpointLabel.Size = New-Object System.Drawing.Size(80, 24)
    $settingsGroup.Controls.Add($endpointLabel)

    $endpointBox = New-Object System.Windows.Forms.TextBox
    $endpointBox.Location = New-Object System.Drawing.Point(102, 61)
    $endpointBox.Size = New-Object System.Drawing.Size(728, 24)
    $endpointBox.Anchor = "Top,Left,Right"
    $endpointBox.Text = $script:AiDefaultEndpoint
    $settingsGroup.Controls.Add($endpointBox)

    $modelLabel = New-Object System.Windows.Forms.Label
    $modelLabel.Text = "Model"
    $modelLabel.Location = New-Object System.Drawing.Point(16, 98)
    $modelLabel.Size = New-Object System.Drawing.Size(80, 24)
    $settingsGroup.Controls.Add($modelLabel)

    $modelBox = New-Object System.Windows.Forms.TextBox
    $modelBox.Location = New-Object System.Drawing.Point(102, 95)
    $modelBox.Size = New-Object System.Drawing.Size(205, 24)
    $modelBox.Text = $script:AiDefaultModel
    $settingsGroup.Controls.Add($modelBox)

    $keyLabel = New-Object System.Windows.Forms.Label
    $keyLabel.Text = "API key"
    $keyLabel.Location = New-Object System.Drawing.Point(330, 98)
    $keyLabel.Size = New-Object System.Drawing.Size(65, 24)
    $settingsGroup.Controls.Add($keyLabel)

    $keyBox = New-Object System.Windows.Forms.TextBox
    $keyBox.Location = New-Object System.Drawing.Point(400, 95)
    $keyBox.Size = New-Object System.Drawing.Size(430, 24)
    $keyBox.Anchor = "Top,Left,Right"
    $keyBox.UseSystemPasswordChar = $true
    $keyBox.Text = $script:AiApiKey
    $settingsGroup.Controls.Add($keyBox)

    $privacyLabel = New-Object System.Windows.Forms.Label
    $privacyLabel.Text = "Approved DOI Azure endpoint only. Chat sends your question plus table counts/sample rows; fresh run findings only if checked."
    $privacyLabel.Location = New-Object System.Drawing.Point(16, 124)
    $privacyLabel.Size = New-Object System.Drawing.Size(812, 18)
    $privacyLabel.Anchor = "Top,Left,Right"
    $privacyLabel.ForeColor = $script:DadaColorMutedText
    $settingsGroup.Controls.Add($privacyLabel)

    $chatGroup = New-Object System.Windows.Forms.GroupBox
    $chatGroup.Text = "Chat"
    $chatGroup.Location = New-Object System.Drawing.Point(14, 258)
    $chatGroup.Size = New-Object System.Drawing.Size(854, 292)
    $chatGroup.Anchor = "Top,Left,Right,Bottom"
    $chatGroup.BackColor = $script:DadaColorSurface
    $dialog.Controls.Add($chatGroup)

    $chatBox = New-Object System.Windows.Forms.TextBox
    $chatBox.Location = New-Object System.Drawing.Point(16, 26)
    $chatBox.Size = New-Object System.Drawing.Size(812, 250)
    $chatBox.Anchor = "Top,Left,Right,Bottom"
    $chatBox.Multiline = $true
    $chatBox.ScrollBars = "Vertical"
    $chatBox.ReadOnly = $true
    $chatBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $chatGroup.Controls.Add($chatBox)

    $messageGroup = New-Object System.Windows.Forms.GroupBox
    $messageGroup.Text = "Message"
    $messageGroup.Location = New-Object System.Drawing.Point(14, 564)
    $messageGroup.Size = New-Object System.Drawing.Size(854, 78)
    $messageGroup.Anchor = "Left,Right,Bottom"
    $messageGroup.BackColor = $script:DadaColorSurface
    $dialog.Controls.Add($messageGroup)

    $questionBox = New-Object System.Windows.Forms.TextBox
    $questionBox.Location = New-Object System.Drawing.Point(16, 25)
    $questionBox.Size = New-Object System.Drawing.Size(680, 38)
    $questionBox.Anchor = "Left,Right,Bottom"
    $questionBox.Multiline = $true
    $questionBox.ScrollBars = "Vertical"
    $messageGroup.Controls.Add($questionBox)

    $sendButton = New-Object System.Windows.Forms.Button
    $sendButton.Text = "Send"
    $sendButton.Location = New-Object System.Drawing.Point(710, 24)
    $sendButton.Size = New-Object System.Drawing.Size(58, 34)
    $sendButton.Anchor = "Right,Bottom"
    $messageGroup.Controls.Add($sendButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(774, 24)
    $closeButton.Size = New-Object System.Drawing.Size(58, 34)
    $closeButton.Anchor = "Right,Bottom"
    $messageGroup.Controls.Add($closeButton)

    $aiDialogToolTip = New-Object System.Windows.Forms.ToolTip
    $aiDialogToolTip.SetToolTip($enabledCheck, "Turn this on only when you want to use AI chat or AIMessage guidance in the export.")
    $aiDialogToolTip.SetToolTip($includeGuideChecksCheck, "Runs fresh checks on a temporary database copy before sending compact findings to the model. Leave off for faster general chat.")
    $aiDialogToolTip.SetToolTip($testConnectionButton, "Checks whether the endpoint, model/deployment name, and API key can reach the model.")
    $aiDialogToolTip.SetToolTip($endpointBox, "Paste the approved base endpoint only, ending at .com or .com/.")
    $aiDialogToolTip.SetToolTip($keyBox, "The key is kept in memory while DADA is open and is not written to the export or run log.")

    Apply-DadaModernTheme -Control $dialog
    $headerPanel.BackColor = $script:DadaColorSurface
    $helperIcon.BackColor = $script:DadaColorSurface
    $helperTitleLabel.ForeColor = $script:DadaColorAccent
    $helperSubtitleLabel.ForeColor = $script:DadaColorMutedText
    $settingsGroup.BackColor = $script:DadaColorSurface
    $chatGroup.BackColor = $script:DadaColorSurface
    $messageGroup.BackColor = $script:DadaColorSurface
    $privacyLabel.ForeColor = $script:DadaColorMutedText
    $chatBox.BackColor = [System.Drawing.Color]::White
    $chatBox.ForeColor = $script:DadaColorText
    $questionBox.BackColor = [System.Drawing.Color]::White
    $questionBox.ForeColor = $script:DadaColorText
    Set-DadaModernButton -Button $testConnectionButton -Style "Secondary"
    Set-DadaModernButton -Button $sendButton -Style "Primary"
    Set-DadaModernButton -Button $closeButton -Style "Secondary"

    $enabledCheck.Add_CheckedChanged({
        $aiEnabledCheck.Checked = $enabledCheck.Checked
    })

    $saveAiHelperSettings = {
        $aiEnabledCheck.Checked = $enabledCheck.Checked
        $script:AiDefaultEndpoint = $endpointBox.Text.Trim()
        $script:AiDefaultModel = $modelBox.Text.Trim()
        $script:AiApiKey = $keyBox.Text
    }

    $dialog.Add_FormClosing({
        & $saveAiHelperSettings
    })

    $closeButton.Add_Click({
        & $saveAiHelperSettings
        $dialog.Close()
    })

    $testConnectionButton.Add_Click({
        & $saveAiHelperSettings
        $testConnectionButton.Enabled = $false
        $sendButton.Enabled = $false
        $dialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $chatBox.AppendText("Testing AI connection...`r`n`r`n")
        try {
            $result = Invoke-AiConnectionTestRequest `
                -Endpoint $endpointBox.Text.Trim() `
                -ApiKey $keyBox.Text `
                -Model $modelBox.Text.Trim()
            $chatBox.AppendText("$result`r`n`r`n")
            [System.Windows.Forms.MessageBox]::Show($result, "AI connection test", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            $message = "AI connection test failed: $($_.Exception.Message)"
            $chatBox.AppendText("$message`r`n`r`n")
            [System.Windows.Forms.MessageBox]::Show($message, "AI connection test failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        finally {
            $dialog.Cursor = [System.Windows.Forms.Cursors]::Default
            $testConnectionButton.Enabled = $true
            $sendButton.Enabled = $true
        }
    })

    $sendButton.Add_Click({
        if (-not $enabledCheck.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Turn on AI helper first.", "AI helper") | Out-Null
            return
        }

        $question = $questionBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($question)) { return }

        & $saveAiHelperSettings
        $chatBox.AppendText("You: $question`r`n`r`n")
        $questionBox.Clear()
        $sendButton.Enabled = $false
        $dialog.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            if ($includeGuideChecksCheck.Checked) {
                $chatBox.AppendText("Reading database context and running fresh checks on a temporary copy...`r`n`r`n")
            }
            else {
                $chatBox.AppendText("Reading database context...`r`n`r`n")
            }
            $context = Get-DatabaseContextForAi -IncludeGuideChecks ([bool]$includeGuideChecksCheck.Checked)
            $answer = Invoke-AiChatRequest `
                -Endpoint $endpointBox.Text.Trim() `
                -ApiKey $keyBox.Text `
                -Model $modelBox.Text.Trim() `
                -Question $question `
                -AuditContext $context `
                -PriorChat $chatBox.Text
            $chatBox.AppendText("AI: $answer`r`n`r`n")
        }
        catch {
            $chatBox.AppendText("AI request failed: $($_.Exception.Message)`r`n`r`n")
        }
        finally {
            $dialog.Cursor = [System.Windows.Forms.Cursors]::Default
            $sendButton.Enabled = $true
        }
    })

    [void]$dialog.ShowDialog($form)
    $dialog.Dispose()
}

Clear-StaleAppTempFiles

if ($env:FOREST_CLEANER_NO_UI -eq "1") {
    return
}

Set-DadaTaskbarIdentity

$script:DadaColorBackground = [System.Drawing.Color]::FromArgb(244, 247, 245)
$script:DadaColorSurface = [System.Drawing.Color]::FromArgb(255, 255, 255)
$script:DadaColorSurfaceMuted = [System.Drawing.Color]::FromArgb(238, 243, 238)
$script:DadaColorBorder = [System.Drawing.Color]::FromArgb(205, 216, 207)
$script:DadaColorText = [System.Drawing.Color]::FromArgb(31, 41, 33)
$script:DadaColorMutedText = [System.Drawing.Color]::FromArgb(89, 103, 91)
$script:DadaColorAccent = [System.Drawing.Color]::FromArgb(43, 93, 68)
$script:DadaColorAccentHover = [System.Drawing.Color]::FromArgb(36, 78, 58)
$script:DadaColorSoftAccent = [System.Drawing.Color]::FromArgb(221, 235, 224)
$script:DadaColorDanger = [System.Drawing.Color]::FromArgb(159, 55, 55)
$script:DadaColorAiSurface = [System.Drawing.Color]::FromArgb(232, 241, 252)
$script:DadaColorAiSurfaceMuted = [System.Drawing.Color]::FromArgb(219, 234, 251)
$script:DadaColorAiAccent = [System.Drawing.Color]::FromArgb(37, 99, 150)

function Set-DadaModernButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [ValidateSet("Primary", "Secondary", "Danger")]
        [string]$Style = "Secondary"
    )

    if ($null -eq $Button) { return }

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
    $Button.FlatAppearance.BorderSize = 1

    switch ($Style) {
        "Primary" {
            $Button.BackColor = $script:DadaColorAccent
            $Button.ForeColor = [System.Drawing.Color]::White
            $Button.FlatAppearance.BorderColor = $script:DadaColorAccent
            $Button.FlatAppearance.MouseOverBackColor = $script:DadaColorAccentHover
            $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(28, 62, 46)
        }
        "Danger" {
            $Button.BackColor = $script:DadaColorDanger
            $Button.ForeColor = [System.Drawing.Color]::White
            $Button.FlatAppearance.BorderColor = $script:DadaColorDanger
            $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(133, 45, 45)
            $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(107, 36, 36)
        }
        default {
            $Button.BackColor = $script:DadaColorSurface
            $Button.ForeColor = $script:DadaColorAccent
            $Button.FlatAppearance.BorderColor = $script:DadaColorBorder
            $Button.FlatAppearance.MouseOverBackColor = $script:DadaColorSoftAccent
            $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(203, 222, 207)
        }
    }
}

function Set-DadaModernInput {
    param([System.Windows.Forms.Control]$Control)

    if ($null -eq $Control) { return }

    $Control.BackColor = [System.Drawing.Color]::White
    $Control.ForeColor = $script:DadaColorText
    if ($Control -is [System.Windows.Forms.TextBoxBase]) {
        $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
}

function Apply-DadaModernTheme {
    param([System.Windows.Forms.Control]$Control)

    if ($null -eq $Control) { return }

    if ($Control -is [System.Windows.Forms.GroupBox]) {
        $Control.BackColor = $script:DadaColorSurface
        $Control.ForeColor = $script:DadaColorText
        $Control.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 12)
    }
    elseif ($Control -is [System.Windows.Forms.Panel]) {
        if ($Control.BackColor.ToArgb() -eq [System.Drawing.SystemColors]::Control.ToArgb()) {
            $Control.BackColor = $script:DadaColorBackground
        }
    }
    elseif ($Control -is [System.Windows.Forms.TabPage]) {
        $Control.BackColor = $script:DadaColorBackground
        $Control.ForeColor = $script:DadaColorText
    }
    elseif ($Control -is [System.Windows.Forms.Label]) {
        $Control.ForeColor = $script:DadaColorText
    }
    elseif (($Control -is [System.Windows.Forms.TextBoxBase]) -or
        ($Control -is [System.Windows.Forms.ComboBox]) -or
        ($Control -is [System.Windows.Forms.NumericUpDown])) {
        Set-DadaModernInput -Control $Control
    }
    elseif ($Control -is [System.Windows.Forms.CheckBox]) {
        $Control.ForeColor = $script:DadaColorText
        $Control.BackColor = [System.Drawing.Color]::Transparent
    }

    foreach ($child in $Control.Controls) {
        Apply-DadaModernTheme -Control $child
    }
}

function Show-DadaSplashScreen {
    if ($env:FOREST_CLEANER_NO_UI -eq "1") { return $null }

    $splash = New-Object System.Windows.Forms.Form
    $splash.Text = "$script:AppName is loading"
    $splash.StartPosition = "CenterScreen"
    $splash.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $splash.Size = New-Object System.Drawing.Size(470, 220)
    $splash.BackColor = $script:DadaColorSurface
    $splash.ShowInTaskbar = $false
    $splash.TopMost = $true
    try {
        if ($null -ne $script:AppIcon) { $splash.Icon = $script:AppIcon }
    }
    catch {
    }

    $outerPanel = New-Object System.Windows.Forms.Panel
    $outerPanel.Dock = "Fill"
    $outerPanel.BackColor = $script:DadaColorSurface
    $outerPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $splash.Controls.Add($outerPanel)

    $iconPicture = New-Object System.Windows.Forms.PictureBox
    $iconPicture.Location = New-Object System.Drawing.Point(30, 42)
    $iconPicture.Size = New-Object System.Drawing.Size(86, 86)
    $iconPicture.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $iconPicture.BackColor = $script:DadaColorSurface
    try {
        if ($null -ne $script:AppIcon) {
            $iconPicture.Image = $script:AppIcon.ToBitmap()
        }
    }
    catch {
    }
    $outerPanel.Controls.Add($iconPicture)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $script:AppName
    $titleLabel.Location = New-Object System.Drawing.Point(136, 38)
    $titleLabel.Size = New-Object System.Drawing.Size(300, 34)
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 17)
    $titleLabel.ForeColor = $script:DadaColorAccent
    $outerPanel.Controls.Add($titleLabel)

    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "Loading forest inventory cleaner..."
    $subtitleLabel.Location = New-Object System.Drawing.Point(138, 82)
    $subtitleLabel.Size = New-Object System.Drawing.Size(300, 24)
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitleLabel.ForeColor = $script:DadaColorText
    $outerPanel.Controls.Add($subtitleLabel)

    $detailLabel = New-Object System.Windows.Forms.Label
    $detailLabel.Text = "Preparing the run workspace and controls."
    $detailLabel.Location = New-Object System.Drawing.Point(138, 112)
    $detailLabel.Size = New-Object System.Drawing.Size(300, 22)
    $detailLabel.ForeColor = $script:DadaColorMutedText
    $outerPanel.Controls.Add($detailLabel)

    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Build $script:AppVersion"
    $versionLabel.Location = New-Object System.Drawing.Point(28, 170)
    $versionLabel.Size = New-Object System.Drawing.Size(410, 20)
    $versionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $versionLabel.ForeColor = $script:DadaColorMutedText
    $outerPanel.Controls.Add($versionLabel)

    [void]$splash.Show()
    [System.Windows.Forms.Application]::DoEvents()
    return $splash
}

function Close-DadaSplashScreen {
    param([System.Windows.Forms.Form]$Splash)

    if ($null -eq $Splash) { return }

    try {
        $Splash.Close()
        $Splash.Dispose()
        [System.Windows.Forms.Application]::DoEvents()
    }
    catch {
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = $script:AppWindowBaseTitle
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1250, 900)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 760)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = $script:DadaColorBackground
$form.ShowInTaskbar = $true
try {
    $script:AppIcon = New-DadaAppIcon
    if ($null -ne $script:AppIcon) { $form.Icon = $script:AppIcon }
}
catch {
    $script:AppIcon = $null
}

$script:StartupSplash = Show-DadaSplashScreen

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = "Bottom"
$statusStrip.BackColor = $script:DadaColorSurface
$statusStrip.ForeColor = $script:DadaColorMutedText
$statusStrip.SizingGrip = $false
$statusStripLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusStripLabel.Text = "Ready"
$statusStripLabel.Spring = $true
$statusStripLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusStripProgressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$statusStripProgressBar.Minimum = 0
$statusStripProgressBar.Maximum = 100
$statusStripProgressBar.Value = 0
$statusStripProgressBar.Size = New-Object System.Drawing.Size(220, 16)
[void]$statusStrip.Items.Add($statusStripLabel)
[void]$statusStrip.Items.Add($statusStripProgressBar)
$form.Controls.Add($statusStrip)

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = "Top"
$topPanel.Height = 126
$topPanel.BackColor = $script:DadaColorSurface
$form.Controls.Add($topPanel)

$headerTitleLabel = New-Object System.Windows.Forms.Label
$headerTitleLabel.Text = $script:AppName
$headerTitleLabel.Location = New-Object System.Drawing.Point(12, 10)
$headerTitleLabel.Size = New-Object System.Drawing.Size(360, 28)
$headerTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14)
$headerTitleLabel.ForeColor = $script:DadaColorAccent
$topPanel.Controls.Add($headerTitleLabel)

$headerSubtitleLabel = New-Object System.Windows.Forms.Label
$headerSubtitleLabel.Text = "Developed by the BIA Division of Forestry, Branch of Inventory and Planning"
$headerSubtitleLabel.Location = New-Object System.Drawing.Point(14, 38)
$headerSubtitleLabel.Size = New-Object System.Drawing.Size(600, 18)
$headerSubtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$headerSubtitleLabel.ForeColor = $script:DadaColorMutedText
$topPanel.Controls.Add($headerSubtitleLabel)

$topDivider = New-Object System.Windows.Forms.Panel
$topDivider.Dock = "Bottom"
$topDivider.Height = 1
$topDivider.BackColor = $script:DadaColorBorder
$topPanel.Controls.Add($topDivider)

$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Access database"
$pathLabel.Location = New-Object System.Drawing.Point(12, 62)
$pathLabel.Size = New-Object System.Drawing.Size(110, 24)
$topPanel.Controls.Add($pathLabel)

$pathBox = New-Object System.Windows.Forms.TextBox
$pathBox.Location = New-Object System.Drawing.Point(125, 59)
$pathBox.Size = New-Object System.Drawing.Size(900, 24)
$pathBox.Anchor = "Top,Left,Right"
$topPanel.Controls.Add($pathBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse"
$browseButton.Location = New-Object System.Drawing.Point(1035, 58)
$browseButton.Size = New-Object System.Drawing.Size(80, 28)
$browseButton.Anchor = "Top,Right"
$topPanel.Controls.Add($browseButton)

$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Text = "Password"
$passwordLabel.Location = New-Object System.Drawing.Point(12, 96)
$passwordLabel.Size = New-Object System.Drawing.Size(110, 24)
$passwordLabel.Visible = $false
$topPanel.Controls.Add($passwordLabel)

$passwordBox = New-Object System.Windows.Forms.TextBox
$passwordBox.Location = New-Object System.Drawing.Point(125, 93)
$passwordBox.Size = New-Object System.Drawing.Size(220, 24)
$passwordBox.UseSystemPasswordChar = $true
$passwordBox.Text = ""
$passwordBox.Visible = $false
$topPanel.Controls.Add($passwordBox)

$tableLabel = New-Object System.Windows.Forms.Label
$tableLabel.Text = "Inventory table"
$tableLabel.Location = New-Object System.Drawing.Point(370, 96)
$tableLabel.Size = New-Object System.Drawing.Size(100, 24)
$tableLabel.Visible = $false
$topPanel.Controls.Add($tableLabel)

$tableCombo = New-Object System.Windows.Forms.ComboBox
$tableCombo.Location = New-Object System.Drawing.Point(475, 93)
$tableCombo.Size = New-Object System.Drawing.Size(360, 24)
$tableCombo.DropDownStyle = "DropDownList"
$tableCombo.Anchor = "Top,Left,Right"
$tableCombo.Visible = $false
$topPanel.Controls.Add($tableCombo)

$buildLabel = New-Object System.Windows.Forms.Label
$buildLabel.Text = "Build $script:AppVersion"
$buildLabel.Location = New-Object System.Drawing.Point(955, 18)
$buildLabel.Size = New-Object System.Drawing.Size(160, 26)
$buildLabel.Anchor = "Top,Right"
$buildLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$buildLabel.BackColor = $script:DadaColorSoftAccent
$buildLabel.ForeColor = $script:DadaColorAccent
$buildLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$topPanel.Controls.Add($buildLabel)

$projectManualLabel = New-Object System.Windows.Forms.Label
$projectManualLabel.Text = "Project manual"
$projectManualLabel.Location = New-Object System.Drawing.Point(12, 96)
$projectManualLabel.Size = New-Object System.Drawing.Size(110, 24)
$topPanel.Controls.Add($projectManualLabel)

$projectManualBox = New-Object System.Windows.Forms.TextBox
$projectManualBox.Location = New-Object System.Drawing.Point(125, 93)
$projectManualBox.Size = New-Object System.Drawing.Size(560, 24)
$projectManualBox.ReadOnly = $true
$projectManualBox.Anchor = "Top,Left,Right"
$topPanel.Controls.Add($projectManualBox)

$projectManualButton = New-Object System.Windows.Forms.Button
$projectManualButton.Text = "Browse"
$projectManualButton.Location = New-Object System.Drawing.Point(695, 92)
$projectManualButton.Size = New-Object System.Drawing.Size(80, 30)
$projectManualButton.Anchor = "Top,Right"
$topPanel.Controls.Add($projectManualButton)

$creditLabel = New-Object System.Windows.Forms.Label
$creditLabel.Text = "Developed by the BIA Division of Forestry, Branch of Forest Inventory and Planning."
$creditLabel.Location = New-Object System.Drawing.Point(125, 92)
$creditLabel.Size = New-Object System.Drawing.Size(780, 14)
$creditLabel.Anchor = "Top,Left,Right"
$creditLabel.ForeColor = $script:DadaColorMutedText
$creditLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$creditLabel.Visible = $false
$topPanel.Controls.Add($creditLabel)

$codeChangeNoteLabel = New-Object System.Windows.Forms.Label
$codeChangeNoteLabel.Text = "Crosswalk note: run the main DADA cleaning after project codes are crosswalked. During a crosswalk, ignore cleaning-style findings caused by old code meanings and treat them as conversion review."
$codeChangeNoteLabel.Location = New-Object System.Drawing.Point(125, 88)
$codeChangeNoteLabel.Size = New-Object System.Drawing.Size(990, 34)
$codeChangeNoteLabel.Anchor = "Top,Left,Right"
$codeChangeNoteLabel.ForeColor = $script:DadaColorMutedText
$codeChangeNoteLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$topPanel.Controls.Add($codeChangeNoteLabel)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear all"
$clearButton.Size = New-Object System.Drawing.Size(100, 34)

$topGuideButton = New-Object System.Windows.Forms.Button
$topGuideButton.Text = "Run"
$topGuideButton.Location = New-Object System.Drawing.Point(475, 96)
$topGuideButton.Size = New-Object System.Drawing.Size(120, 30)
$topGuideButton.Visible = $false
$topPanel.Controls.Add($topGuideButton)

$topAiEnabledCheck = New-Object System.Windows.Forms.CheckBox
$topAiEnabledCheck.Text = "AI helper on"
$topAiEnabledCheck.Location = New-Object System.Drawing.Point(615, 100)
$topAiEnabledCheck.Size = New-Object System.Drawing.Size(115, 24)
$topAiEnabledCheck.Visible = $false
$topPanel.Controls.Add($topAiEnabledCheck)

$topAiHelperButton = New-Object System.Windows.Forms.Button
$topAiHelperButton.Text = "Open AI helper"
$topAiHelperButton.Location = New-Object System.Drawing.Point(740, 96)
$topAiHelperButton.Size = New-Object System.Drawing.Size(135, 30)
$topAiHelperButton.Visible = $false
$topPanel.Controls.Add($topAiHelperButton)

$topCancelRunButton = New-Object System.Windows.Forms.Button
$topCancelRunButton.Text = "Cancel run"
$topCancelRunButton.Location = New-Object System.Drawing.Point(885, 96)
$topCancelRunButton.Size = New-Object System.Drawing.Size(105, 30)
$topCancelRunButton.Visible = $false
$topCancelRunButton.Enabled = $false
$topPanel.Controls.Add($topCancelRunButton)

$mappingGroup = New-Object System.Windows.Forms.GroupBox
$mappingGroup.Text = "Field mapping"
$mappingGroup.Location = New-Object System.Drawing.Point(12, 118)
$mappingGroup.Size = New-Object System.Drawing.Size(1118, 105)
$mappingGroup.Anchor = "Top,Left,Right"
$mappingGroup.Visible = $false
$topPanel.Controls.Add($mappingGroup)

function Add-MappingPair {
    param(
        [System.Windows.Forms.GroupBox]$Parent,
        [string]$Label,
        [int]$X,
        [int]$Y
    )

    $labelControl = New-Object System.Windows.Forms.Label
    $labelControl.Text = $Label
    $labelControl.Location = New-Object System.Drawing.Point -ArgumentList $X, ($Y + 4)
    $labelControl.Size = New-Object System.Drawing.Size(80, 24)
    $Parent.Controls.Add($labelControl)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Location = New-Object System.Drawing.Point -ArgumentList ($X + 85), $Y
    $combo.Size = New-Object System.Drawing.Size(205, 24)
    $combo.DropDownStyle = "DropDownList"
    $Parent.Controls.Add($combo)

    return $combo
}

$speciesCombo = Add-MappingPair -Parent $mappingGroup -Label "Species" -X 14 -Y 28
$dbhCombo = Add-MappingPair -Parent $mappingGroup -Label "IDBH" -X 335 -Y 28
$heightCombo = Add-MappingPair -Parent $mappingGroup -Label "Height" -X 655 -Y 28
$plotCombo = Add-MappingPair -Parent $mappingGroup -Label "Plot/key" -X 14 -Y 65
$treeCombo = Add-MappingPair -Parent $mappingGroup -Label "Tree/regen" -X 335 -Y 65

$mainTabControl = New-Object System.Windows.Forms.TabControl
$mainTabControl.Dock = "Fill"
$mainTabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$mainTabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$form.Controls.Add($mainTabControl)
$form.Controls.SetChildIndex($mainTabControl, 0)

$runTabPage = New-Object System.Windows.Forms.TabPage
$runTabPage.Text = "Run"
$runTabPage.BackColor = $script:DadaColorBackground
[void]$mainTabControl.TabPages.Add($runTabPage)

$aiTabPage = New-Object System.Windows.Forms.TabPage
$aiTabPage.Text = "AI guidance"
$aiTabPage.BackColor = $script:DadaColorAiSurface
[void]$mainTabControl.TabPages.Add($aiTabPage)

$mainTabControl.Add_DrawItem({
    param($sender, $eventArgs)

    $tabPage = $sender.TabPages[$eventArgs.Index]
    $isSelected = (($eventArgs.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected)
    $isAiTab = ([string]$tabPage.Text -eq "AI guidance")

    if ($isAiTab) {
        $backColor = if ($isSelected) { $script:DadaColorAiAccent } else { $script:DadaColorAiSurfaceMuted }
        $foreColor = if ($isSelected) { [System.Drawing.Color]::White } else { $script:DadaColorAiAccent }
    }
    else {
        $backColor = if ($isSelected) { $script:DadaColorSurface } else { $script:DadaColorBackground }
        $foreColor = if ($isSelected) { $script:DadaColorAccent } else { $script:DadaColorText }
    }

    $backgroundBrush = New-Object System.Drawing.SolidBrush -ArgumentList $backColor
    try {
        $eventArgs.Graphics.FillRectangle($backgroundBrush, $eventArgs.Bounds)
    }
    finally {
        $backgroundBrush.Dispose()
    }

    $textFlags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis
    [System.Windows.Forms.TextRenderer]::DrawText($eventArgs.Graphics, [string]$tabPage.Text, $sender.Font, $eventArgs.Bounds, $foreColor, $textFlags)
})

$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = "Fill"
$mainPanel.AutoScroll = $true
$mainPanel.AutoScrollMinSize = New-Object System.Drawing.Size(1228, 1280)
$mainPanel.BackColor = $script:DadaColorBackground
$runTabPage.Controls.Add($mainPanel)

$aiPanel = New-Object System.Windows.Forms.Panel
$aiPanel.Dock = "Fill"
$aiPanel.AutoScroll = $true
$aiPanel.AutoScrollMinSize = New-Object System.Drawing.Size(0, 340)
$aiPanel.BackColor = $script:DadaColorAiSurface
$aiTabPage.Controls.Add($aiPanel)

$optionsGroup = New-Object System.Windows.Forms.GroupBox
$optionsGroup.Text = "Run options"
$optionsGroup.Location = New-Object System.Drawing.Point(16, 14)
$optionsGroup.Size = New-Object System.Drawing.Size(1192, 750)
$optionsGroup.Anchor = "Top,Left"
$optionsGroup.Height = 750
$optionsGroup.BackColor = $script:DadaColorSurface
$mainPanel.Controls.Add($optionsGroup)

$trimTextCheck = New-Object System.Windows.Forms.CheckBox
$trimTextCheck.Text = "Trim leading and trailing spaces from text fields"
$trimTextCheck.Location = New-Object System.Drawing.Point(15, 30)
$trimTextCheck.Size = New-Object System.Drawing.Size(360, 24)
$trimTextCheck.Checked = $true
$trimTextCheck.Visible = $false
$optionsGroup.Controls.Add($trimTextCheck)

$blankNullCheck = New-Object System.Windows.Forms.CheckBox
$blankNullCheck.Text = "Convert blank text values to Null"
$blankNullCheck.Location = New-Object System.Drawing.Point(15, 58)
$blankNullCheck.Size = New-Object System.Drawing.Size(280, 24)
$blankNullCheck.Checked = $true
$blankNullCheck.Visible = $false
$optionsGroup.Controls.Add($blankNullCheck)

$speciesUpperCheck = New-Object System.Windows.Forms.CheckBox
$speciesUpperCheck.Text = "Uppercase species codes"
$speciesUpperCheck.Location = New-Object System.Drawing.Point(15, 86)
$speciesUpperCheck.Size = New-Object System.Drawing.Size(260, 24)
$speciesUpperCheck.Checked = $true
$speciesUpperCheck.Visible = $false
$optionsGroup.Controls.Add($speciesUpperCheck)

$rangeCheck = New-Object System.Windows.Forms.CheckBox
$rangeCheck.Text = "Run range checks"
$rangeCheck.Location = New-Object System.Drawing.Point(15, 30)
$rangeCheck.Size = New-Object System.Drawing.Size(180, 24)
$rangeCheck.Checked = $true
$optionsGroup.Controls.Add($rangeCheck)

$duplicateCheck = New-Object System.Windows.Forms.CheckBox
$duplicateCheck.Text = "Flag duplicate selected keys"
$duplicateCheck.Location = New-Object System.Drawing.Point(15, 142)
$duplicateCheck.Size = New-Object System.Drawing.Size(260, 24)
$duplicateCheck.Checked = $true
$duplicateCheck.Visible = $false
$optionsGroup.Controls.Add($duplicateCheck)

$splitHistoryColumnsCheck = New-Object System.Windows.Forms.CheckBox
$splitHistoryColumnsCheck.Text = "Split DBH/height/TreeHistory into columns"
$splitHistoryColumnsCheck.Location = New-Object System.Drawing.Point(775, 96)
$splitHistoryColumnsCheck.Size = New-Object System.Drawing.Size(335, 24)
$splitHistoryColumnsCheck.Checked = $true
$splitHistoryColumnsCheck.Visible = $false
$optionsGroup.Controls.Add($splitHistoryColumnsCheck)

$dbhMaxLabel = New-Object System.Windows.Forms.Label
$dbhMaxLabel.Text = "Max IDBH"
$dbhMaxLabel.Location = New-Object System.Drawing.Point(15, 68)
$dbhMaxLabel.Size = New-Object System.Drawing.Size(105, 24)
$optionsGroup.Controls.Add($dbhMaxLabel)

$dbhMax = New-Object System.Windows.Forms.NumericUpDown
$dbhMax.Location = New-Object System.Drawing.Point(130, 65)
$dbhMax.Minimum = 1
$dbhMax.Maximum = 3500
$dbhMax.DecimalPlaces = 0
$dbhMax.Increment = 1
$dbhMax.Value = 500
$dbhMax.Size = New-Object System.Drawing.Size(75, 24)
$optionsGroup.Controls.Add($dbhMax)

$heightMaxLabel = New-Object System.Windows.Forms.Label
$heightMaxLabel.Text = "Max height"
$heightMaxLabel.Location = New-Object System.Drawing.Point(15, 250)
$heightMaxLabel.Size = New-Object System.Drawing.Size(105, 24)
$optionsGroup.Controls.Add($heightMaxLabel)

$heightMax = New-Object System.Windows.Forms.NumericUpDown
$heightMax.Location = New-Object System.Drawing.Point(130, 247)
$heightMax.Minimum = 1
$heightMax.Maximum = 1000
$heightMax.DecimalPlaces = 0
$heightMax.Increment = 1
$heightMax.Value = 150
$heightMax.Size = New-Object System.Drawing.Size(75, 24)
$optionsGroup.Controls.Add($heightMax)

$dbhGrowthLabel = New-Object System.Windows.Forms.Label
$dbhGrowthLabel.Text = "IDBH jump"
$dbhGrowthLabel.Location = New-Object System.Drawing.Point(230, 68)
$dbhGrowthLabel.Size = New-Object System.Drawing.Size(95, 24)
$optionsGroup.Controls.Add($dbhGrowthLabel)

$dbhGrowthMax = New-Object System.Windows.Forms.NumericUpDown
$dbhGrowthMax.Location = New-Object System.Drawing.Point(340, 65)
$dbhGrowthMax.Minimum = 1
$dbhGrowthMax.Maximum = 1000
$dbhGrowthMax.DecimalPlaces = 0
$dbhGrowthMax.Increment = 1
$dbhGrowthMax.Value = 100
$dbhGrowthMax.Size = New-Object System.Drawing.Size(75, 24)
$optionsGroup.Controls.Add($dbhGrowthMax)

$heightGrowthLabel = New-Object System.Windows.Forms.Label
$heightGrowthLabel.Text = "Height jump"
$heightGrowthLabel.Location = New-Object System.Drawing.Point(230, 250)
$heightGrowthLabel.Size = New-Object System.Drawing.Size(95, 24)
$optionsGroup.Controls.Add($heightGrowthLabel)

$heightGrowthMax = New-Object System.Windows.Forms.NumericUpDown
$heightGrowthMax.Location = New-Object System.Drawing.Point(340, 247)
$heightGrowthMax.Minimum = 1
$heightGrowthMax.Maximum = 1000
$heightGrowthMax.DecimalPlaces = 0
$heightGrowthMax.Increment = 1
$heightGrowthMax.Value = 20
$heightGrowthMax.Size = New-Object System.Drawing.Size(75, 24)
$optionsGroup.Controls.Add($heightGrowthMax)

$periodScopeCheck = New-Object System.Windows.Forms.CheckBox
$periodScopeCheck.Text = "Limit normal checks to measurement period(s)"
$periodScopeCheck.Location = New-Object System.Drawing.Point(230, 30)
$periodScopeCheck.Size = New-Object System.Drawing.Size(315, 24)
$periodScopeCheck.Checked = $true
$optionsGroup.Controls.Add($periodScopeCheck)

$singleMeasurementProjectCheck = New-Object System.Windows.Forms.CheckBox
$singleMeasurementProjectCheck.Text = "Single measurement project"
$singleMeasurementProjectCheck.Location = New-Object System.Drawing.Point(555, 30)
$singleMeasurementProjectCheck.Size = New-Object System.Drawing.Size(210, 24)
$singleMeasurementProjectCheck.Checked = $false
$optionsGroup.Controls.Add($singleMeasurementProjectCheck)

$currentPeriodLabel = New-Object System.Windows.Forms.Label
$currentPeriodLabel.Text = "Current measurement period #"
$currentPeriodLabel.Location = New-Object System.Drawing.Point(455, 68)
$currentPeriodLabel.Size = New-Object System.Drawing.Size(205, 24)
$optionsGroup.Controls.Add($currentPeriodLabel)

$currentPeriodBox = New-Object System.Windows.Forms.NumericUpDown
$currentPeriodBox.Location = New-Object System.Drawing.Point(665, 65)
$currentPeriodBox.Minimum = 0
$currentPeriodBox.Maximum = 999
$currentPeriodBox.DecimalPlaces = 0
$currentPeriodBox.Increment = 1
$currentPeriodBox.Value = 0
$currentPeriodBox.Size = New-Object System.Drawing.Size(65, 24)
$optionsGroup.Controls.Add($currentPeriodBox)

$pastPeriodLabel = New-Object System.Windows.Forms.Label
$pastPeriodLabel.Text = "Previous measurement period #"
$pastPeriodLabel.Location = New-Object System.Drawing.Point(750, 68)
$pastPeriodLabel.Size = New-Object System.Drawing.Size(210, 24)
$optionsGroup.Controls.Add($pastPeriodLabel)

$pastPeriodBox = New-Object System.Windows.Forms.NumericUpDown
$pastPeriodBox.Location = New-Object System.Drawing.Point(965, 65)
$pastPeriodBox.Minimum = 0
$pastPeriodBox.Maximum = 999
$pastPeriodBox.DecimalPlaces = 0
$pastPeriodBox.Increment = 1
$pastPeriodBox.Value = 0
$pastPeriodBox.Size = New-Object System.Drawing.Size(65, 24)
$optionsGroup.Controls.Add($pastPeriodBox)

$periodScopeNoteLabel = New-Object System.Windows.Forms.Label
$periodScopeNoteLabel.Text = "TreeHistory checks all periods."
$periodScopeNoteLabel.Location = New-Object System.Drawing.Point(775, 32)
$periodScopeNoteLabel.Size = New-Object System.Drawing.Size(335, 22)
$periodScopeNoteLabel.Anchor = "Top,Left,Right"
$periodScopeNoteLabel.ForeColor = $script:DadaColorMutedText
$optionsGroup.Controls.Add($periodScopeNoteLabel)

$stemCountLabel = New-Object System.Windows.Forms.Label
$stemCountLabel.Text = "Max regen stems"
$stemCountLabel.Location = New-Object System.Drawing.Point(15, 512)
$stemCountLabel.Size = New-Object System.Drawing.Size(135, 24)
$optionsGroup.Controls.Add($stemCountLabel)

$stemCountMax = New-Object System.Windows.Forms.NumericUpDown
$stemCountMax.Location = New-Object System.Drawing.Point(160, 509)
$stemCountMax.Minimum = 1
$stemCountMax.Maximum = 32767
$stemCountMax.DecimalPlaces = 0
$stemCountMax.Value = 100
$stemCountMax.Size = New-Object System.Drawing.Size(75, 24)
$optionsGroup.Controls.Add($stemCountMax)

$treeStemCountCheck = New-Object System.Windows.Forms.CheckBox
$treeStemCountCheck.Text = "Check tree StemCount"
$treeStemCountCheck.Location = New-Object System.Drawing.Point(15, 186)
$treeStemCountCheck.Size = New-Object System.Drawing.Size(220, 24)
$treeStemCountCheck.Checked = $false
$optionsGroup.Controls.Add($treeStemCountCheck)

$newMortalityIdbhCheck = New-Object System.Windows.Forms.CheckBox
$newMortalityIdbhCheck.Text = "Require DBH for new mortality (TreeHistory 2/3)"
$newMortalityIdbhCheck.Location = New-Object System.Drawing.Point(15, 100)
$newMortalityIdbhCheck.Size = New-Object System.Drawing.Size(330, 24)
$newMortalityIdbhCheck.Checked = $false
$optionsGroup.Controls.Add($newMortalityIdbhCheck)

$oldMortalityIdbhCheck = New-Object System.Windows.Forms.CheckBox
$oldMortalityIdbhCheck.Text = "Require DBH for old mortality (TreeHistory 7)"
$oldMortalityIdbhCheck.Location = New-Object System.Drawing.Point(360, 100)
$oldMortalityIdbhCheck.Size = New-Object System.Drawing.Size(320, 24)
$oldMortalityIdbhCheck.Checked = $false
$optionsGroup.Controls.Add($oldMortalityIdbhCheck)

$newMortalityHeightCheck = New-Object System.Windows.Forms.CheckBox
$newMortalityHeightCheck.Text = "Require height for new mortality (TreeHistory 2/3)"
$newMortalityHeightCheck.Location = New-Object System.Drawing.Point(15, 276)
$newMortalityHeightCheck.Size = New-Object System.Drawing.Size(345, 24)
$newMortalityHeightCheck.Checked = $false
$optionsGroup.Controls.Add($newMortalityHeightCheck)

$oldMortalityHeightCheck = New-Object System.Windows.Forms.CheckBox
$oldMortalityHeightCheck.Text = "Require height for old mortality (TreeHistory 7)"
$oldMortalityHeightCheck.Location = New-Object System.Drawing.Point(360, 276)
$oldMortalityHeightCheck.Size = New-Object System.Drawing.Size(330, 24)
$oldMortalityHeightCheck.Checked = $false
$optionsGroup.Controls.Add($oldMortalityHeightCheck)

$problem127HeightCheck = New-Object System.Windows.Forms.CheckBox
$problem127HeightCheck.Text = "Flag timber height: broken top (127/74)"
$problem127HeightCheck.Location = New-Object System.Drawing.Point(15, 304)
$problem127HeightCheck.Size = New-Object System.Drawing.Size(360, 24)
$problem127HeightCheck.Checked = $true
$optionsGroup.Controls.Add($problem127HeightCheck)

$problem128HeightCheck = New-Object System.Windows.Forms.CheckBox
$problem128HeightCheck.Text = "Flag timber height: dead top (128/75)"
$problem128HeightCheck.Location = New-Object System.Drawing.Point(385, 304)
$problem128HeightCheck.Size = New-Object System.Drawing.Size(290, 24)
$problem128HeightCheck.Checked = $true
$optionsGroup.Controls.Add($problem128HeightCheck)

$problem123HeightCheck = New-Object System.Windows.Forms.CheckBox
$problem123HeightCheck.Text = "Flag timber height: lean > 15 degrees (123/72)"
$problem123HeightCheck.Location = New-Object System.Drawing.Point(690, 304)
$problem123HeightCheck.Size = New-Object System.Drawing.Size(395, 24)
$problem123HeightCheck.Checked = $true
$optionsGroup.Controls.Add($problem123HeightCheck)

$woodlandSpeciesCodesLabel = New-Object System.Windows.Forms.Label
$woodlandSpeciesCodesLabel.Text = "Woodland species to exclude from shrinking diameters"
$woodlandSpeciesCodesLabel.Location = New-Object System.Drawing.Point(230, 132)
$woodlandSpeciesCodesLabel.Size = New-Object System.Drawing.Size(305, 24)
$optionsGroup.Controls.Add($woodlandSpeciesCodesLabel)

$woodlandSpeciesCodesBox = New-Object System.Windows.Forms.TextBox
$woodlandSpeciesCodesBox.Location = New-Object System.Drawing.Point(540, 129)
$woodlandSpeciesCodesBox.Size = New-Object System.Drawing.Size(570, 24)
$woodlandSpeciesCodesBox.Anchor = "Top,Left"
$woodlandSpeciesCodesBox.Text = $script:DefaultWoodlandSpeciesCodes
$optionsGroup.Controls.Add($woodlandSpeciesCodesBox)

$optionsToolTip = New-Object System.Windows.Forms.ToolTip
$optionsToolTip.SetToolTip($clearButton, "Clears the selected database, project manual, woodland species codes, and minor-plot rules from this app window.")
$optionsToolTip.SetToolTip($woodlandSpeciesCodesBox, "Enter the exact SpeciesCode values used in the database. Separate codes with commas, spaces, semicolons, or pipes.")
$optionsToolTip.SetToolTip($periodScopeCheck, "When checked, most cleaning checks run only on the entered measurement period numbers. TreeHistory transition checks always use all periods.")
$optionsToolTip.SetToolTip($singleMeasurementProjectCheck, "Use this for first-measurement projects with no previous measurement period. DADA will ignore the Previous measurement period box.")
$optionsToolTip.SetToolTip($splitHistoryColumnsCheck, "Optional export layout. Adds separate columns for parsed DBH comparison values and TreeHistory by period when those values are available.")
$optionsToolTip.SetToolTip($currentPeriodBox, "Current measurement period number. DADA fills this from the highest PeriodNumber when you connect.")
$optionsToolTip.SetToolTip($pastPeriodBox, "Previous measurement period number. DADA fills this from the next lower PeriodNumber when you connect. Disabled for single measurement projects.")
$optionsToolTip.SetToolTip($treeStemCountCheck, "Optional. Checks TreeMeasurements.StemCount. Timber species may be blank; woodland/non-timber species entered in the woodland species box require StemCount.")
$optionsToolTip.SetToolTip($newMortalityIdbhCheck, "Optional. When checked, blank TreeMeasurements.IDBH is flagged for TreeHistory 2 and 3.")
$optionsToolTip.SetToolTip($oldMortalityIdbhCheck, "Optional. When checked, blank TreeMeasurements.IDBH is flagged for TreeHistory 7.")
$optionsToolTip.SetToolTip($newMortalityHeightCheck, "Optional. When checked, blank TreeMeasurements.TotalHeight is flagged for TreeHistory 2 and 3.")
$optionsToolTip.SetToolTip($oldMortalityHeightCheck, "Optional. When checked, blank TreeMeasurements.TotalHeight is flagged for TreeHistory 7.")
$optionsToolTip.SetToolTip($problem127HeightCheck, "Optional. When checked, entered TotalHeight is flagged for timber trees with Problem1 or Problem2 code 127 or old legacy code 74 (broken/missing top). Blank height is treated as correct.")
$optionsToolTip.SetToolTip($problem128HeightCheck, "Optional. When checked, entered TotalHeight is flagged for timber trees with Problem1 or Problem2 code 128 or old legacy code 75 (dead top). Blank height is treated as correct.")
$optionsToolTip.SetToolTip($problem123HeightCheck, "Optional. When checked, entered TotalHeight is flagged for timber trees with Problem1 or Problem2 code 123 or old legacy code 72 (lean > 15 degrees). Woodland species entered above are skipped.")

function Update-PeriodScopeControlState {
    try {
        $scopeEnabled = ($null -ne $periodScopeCheck -and $periodScopeCheck.Checked)
        $singleMeasurement = ($null -ne $singleMeasurementProjectCheck -and $singleMeasurementProjectCheck.Checked)
        if ($singleMeasurement) {
            if ($null -ne $currentPeriodBox) { $currentPeriodBox.Value = 1 }
            if ($null -ne $pastPeriodBox) { $pastPeriodBox.Value = 0 }
        }

        if ($null -ne $singleMeasurementProjectCheck) {
            $singleMeasurementProjectCheck.Enabled = $scopeEnabled
        }

        $currentEnabled = ($scopeEnabled -and -not $singleMeasurement)
        foreach ($control in @($currentPeriodLabel, $currentPeriodBox)) {
            if ($null -ne $control) { $control.Enabled = $currentEnabled }
        }

        $pastEnabled = ($scopeEnabled -and -not $singleMeasurement)
        foreach ($control in @($pastPeriodLabel, $pastPeriodBox)) {
            if ($null -ne $control) { $control.Enabled = $pastEnabled }
        }

        foreach ($control in @($periodScopeNoteLabel)) {
            if ($null -ne $control) { $control.Enabled = $scopeEnabled }
        }
    }
    catch {
    }
}

$periodScopeCheck.Add_CheckedChanged({
    Update-PeriodScopeControlState
})

$singleMeasurementProjectCheck.Add_CheckedChanged({
    Update-PeriodScopeControlState
})

Update-PeriodScopeControlState

$woodlandSpeciesExampleLabel = New-Object System.Windows.Forms.Label
$woodlandSpeciesExampleLabel.Text = "Example format: 65, 66, 68 - use exact numeric SpeciesCode values from Trees."
$woodlandSpeciesExampleLabel.Location = New-Object System.Drawing.Point(540, 155)
$woodlandSpeciesExampleLabel.Size = New-Object System.Drawing.Size(570, 22)
$woodlandSpeciesExampleLabel.Anchor = "Top,Left"
$woodlandSpeciesExampleLabel.ForeColor = $script:DadaColorMutedText
$optionsGroup.Controls.Add($woodlandSpeciesExampleLabel)

$heightProtocolSectionLabel = New-Object System.Windows.Forms.Label
$heightProtocolSectionLabel.Text = "Total Height Protocol"
$heightProtocolSectionLabel.Location = New-Object System.Drawing.Point(15, 220)
$heightProtocolSectionLabel.Size = New-Object System.Drawing.Size(220, 22)
$heightProtocolSectionLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$heightProtocolSectionLabel.ForeColor = $script:DadaColorAccent
$optionsGroup.Controls.Add($heightProtocolSectionLabel)

$heightProtocolLabel = New-Object System.Windows.Forms.Label
$heightProtocolLabel.Text = "Height collection rule"
$heightProtocolLabel.Location = New-Object System.Drawing.Point(15, 336)
$heightProtocolLabel.Size = New-Object System.Drawing.Size(145, 24)
$optionsGroup.Controls.Add($heightProtocolLabel)

$heightProtocolCombo = New-Object System.Windows.Forms.ComboBox
$heightProtocolCombo.Location = New-Object System.Drawing.Point(170, 333)
$heightProtocolCombo.Size = New-Object System.Drawing.Size(320, 24)
$heightProtocolCombo.DropDownStyle = "DropDownList"
[void]$heightProtocolCombo.Items.Add("Normal required-height rules")
[void]$heightProtocolCombo.Items.Add("100% live-tree heights")
[void]$heightProtocolCombo.Items.Add("Subsample by species and 2-inch IDBH class")
$heightProtocolCombo.SelectedIndex = 0
$optionsGroup.Controls.Add($heightProtocolCombo)

$heightSubsampleMinimumLabel = New-Object System.Windows.Forms.Label
$heightSubsampleMinimumLabel.Text = "Min count (subsample)"
$heightSubsampleMinimumLabel.Location = New-Object System.Drawing.Point(515, 336)
$heightSubsampleMinimumLabel.Size = New-Object System.Drawing.Size(145, 24)
$optionsGroup.Controls.Add($heightSubsampleMinimumLabel)

$heightSubsampleMinimumBox = New-Object System.Windows.Forms.NumericUpDown
$heightSubsampleMinimumBox.Location = New-Object System.Drawing.Point(665, 333)
$heightSubsampleMinimumBox.Minimum = 1
$heightSubsampleMinimumBox.Maximum = 50
$heightSubsampleMinimumBox.DecimalPlaces = 0
$heightSubsampleMinimumBox.Value = 2
$heightSubsampleMinimumBox.Size = New-Object System.Drawing.Size(65, 24)
$optionsGroup.Controls.Add($heightSubsampleMinimumBox)

$heightSubsampleMinIdbhLabel = New-Object System.Windows.Forms.Label
$heightSubsampleMinIdbhLabel.Text = "Min eligible IDBH"
$heightSubsampleMinIdbhLabel.Location = New-Object System.Drawing.Point(755, 336)
$heightSubsampleMinIdbhLabel.Size = New-Object System.Drawing.Size(125, 24)
$optionsGroup.Controls.Add($heightSubsampleMinIdbhLabel)

$heightSubsampleMinIdbhBox = New-Object System.Windows.Forms.NumericUpDown
$heightSubsampleMinIdbhBox.Location = New-Object System.Drawing.Point(885, 333)
$heightSubsampleMinIdbhBox.Minimum = 0
$heightSubsampleMinIdbhBox.Maximum = 3500
$heightSubsampleMinIdbhBox.DecimalPlaces = 0
$heightSubsampleMinIdbhBox.Value = 50
$heightSubsampleMinIdbhBox.Size = New-Object System.Drawing.Size(75, 24)
$optionsGroup.Controls.Add($heightSubsampleMinIdbhBox)

$heightSubsampleAllAtOrAboveCheck = New-Object System.Windows.Forms.CheckBox
$heightSubsampleAllAtOrAboveCheck.Text = "Require all heights at/above IDBH"
$heightSubsampleAllAtOrAboveCheck.Location = New-Object System.Drawing.Point(15, 364)
$heightSubsampleAllAtOrAboveCheck.Size = New-Object System.Drawing.Size(255, 24)
$heightSubsampleAllAtOrAboveCheck.Checked = $false
$optionsGroup.Controls.Add($heightSubsampleAllAtOrAboveCheck)

$heightSubsampleAllAtOrAboveBox = New-Object System.Windows.Forms.NumericUpDown
$heightSubsampleAllAtOrAboveBox.Location = New-Object System.Drawing.Point(275, 363)
$heightSubsampleAllAtOrAboveBox.Minimum = 0
$heightSubsampleAllAtOrAboveBox.Maximum = 3500
$heightSubsampleAllAtOrAboveBox.DecimalPlaces = 0
$heightSubsampleAllAtOrAboveBox.Value = 170
$heightSubsampleAllAtOrAboveBox.Size = New-Object System.Drawing.Size(75, 24)
$heightSubsampleAllAtOrAboveBox.Enabled = $false
$optionsGroup.Controls.Add($heightSubsampleAllAtOrAboveBox)
$heightSubsampleAllAtOrAboveCheck.Add_CheckedChanged({
    Update-HeightProtocolControlState
})

$heightRareSpeciesCodesLabel = New-Object System.Windows.Forms.Label
$heightRareSpeciesCodesLabel.Text = "Rare species requiring 100% live-tree height"
$heightRareSpeciesCodesLabel.Location = New-Object System.Drawing.Point(380, 364)
$heightRareSpeciesCodesLabel.Size = New-Object System.Drawing.Size(285, 24)
$optionsGroup.Controls.Add($heightRareSpeciesCodesLabel)

$heightRareSpeciesCodesBox = New-Object System.Windows.Forms.TextBox
$heightRareSpeciesCodesBox.Location = New-Object System.Drawing.Point(670, 363)
$heightRareSpeciesCodesBox.Size = New-Object System.Drawing.Size(440, 24)
$heightRareSpeciesCodesBox.Anchor = "Top,Left"
$optionsGroup.Controls.Add($heightRareSpeciesCodesBox)

function Update-HeightProtocolControlState {
    try {
        $subsampleSelected = $false
        if ($null -ne $heightProtocolCombo -and $null -ne $heightProtocolCombo.SelectedItem) {
            $subsampleSelected = ([string]$heightProtocolCombo.SelectedItem).StartsWith("Subsample", [System.StringComparison]::OrdinalIgnoreCase)
        }

        foreach ($control in @(
            $heightSubsampleMinimumLabel,
            $heightSubsampleMinimumBox,
            $heightSubsampleMinIdbhLabel,
            $heightSubsampleMinIdbhBox,
            $heightSubsampleAllAtOrAboveCheck
        )) {
            if ($null -ne $control) { $control.Enabled = $subsampleSelected }
        }

        if ($null -ne $heightSubsampleAllAtOrAboveBox) {
            $heightSubsampleAllAtOrAboveBox.Enabled = ($subsampleSelected -and $null -ne $heightSubsampleAllAtOrAboveCheck -and $heightSubsampleAllAtOrAboveCheck.Checked)
        }
    }
    catch {
    }
}

$heightProtocolAllOptionsToolTipText = "Normal required-height rules: checks height only where selected TreeHistory, mortality, minor-plot, rare-species, and problem-code rules say height is required.`r`n`r`n100% live-tree heights: every live tree with TreeHistory 0, 5, or 10 requires height, except selected timber problem-code no-height cases.`r`n`r`nSubsample mode: checks grouped plot/species/2-inch IDBH class minimums instead of flagging every blank live-tree height."
$heightProtocolSelectedToolTips = @{
    "Normal required-height rules" = "Normal required-height rules: DADA checks height only where the selected TreeHistory, mortality, minor-plot, rare-species, and problem-code rules say height is required."
    "100% live-tree heights" = "100% live-tree heights: every live tree with TreeHistory 0, 5, or 10 requires TotalHeight, except selected timber problem-code no-height cases."
    "Subsample by species and 2-inch IDBH class" = "Subsample mode: DADA checks grouped plot/species/2-inch IDBH class height minimums instead of flagging every blank live-tree height."
}
$setHeightProtocolToolTip = {
    $selectedProtocol = if ($null -ne $heightProtocolCombo -and $null -ne $heightProtocolCombo.SelectedItem) { [string]$heightProtocolCombo.SelectedItem } else { "Normal required-height rules" }
    $selectedToolTipText = if ($heightProtocolSelectedToolTips.ContainsKey($selectedProtocol)) { $heightProtocolSelectedToolTips[$selectedProtocol] } else { $heightProtocolAllOptionsToolTipText }
    $optionsToolTip.SetToolTip($heightProtocolSectionLabel, $heightProtocolAllOptionsToolTipText)
    $optionsToolTip.SetToolTip($heightProtocolLabel, $heightProtocolAllOptionsToolTipText)
    $optionsToolTip.SetToolTip($heightProtocolCombo, $selectedToolTipText)
}
& $setHeightProtocolToolTip
$heightProtocolCombo.Add_SelectedIndexChanged({
    & $setHeightProtocolToolTip
    Update-HeightProtocolControlState
})
$heightProtocolCombo.Add_DropDown({
    $optionsToolTip.Show($heightProtocolAllOptionsToolTipText, $heightProtocolCombo, 0, ($heightProtocolCombo.Height + 4), 9000)
})
$heightProtocolCombo.Add_DropDownClosed({
    $optionsToolTip.Hide($heightProtocolCombo)
    & $setHeightProtocolToolTip
    Update-HeightProtocolControlState
})
$optionsToolTip.SetToolTip($heightSubsampleMinimumBox, "Used in subsample mode. Minimum number of height measurements required for each plot/species/2-inch IDBH class.")
$optionsToolTip.SetToolTip($heightSubsampleMinIdbhBox, "Used in subsample mode. Minimum recorded IDBH for a tree to be eligible. Example: 5.0 inches is recorded as 50.")
$optionsToolTip.SetToolTip($heightSubsampleAllAtOrAboveCheck, "Used in subsample mode. When checked, all eligible trees at or above the entered IDBH must have height.")
$optionsToolTip.SetToolTip($heightSubsampleAllAtOrAboveBox, "Used in subsample mode. Example: 17.0 inches is recorded as 170.")
$optionsToolTip.SetToolTip($heightRareSpeciesCodesBox, "SpeciesCode values that always require live-tree TotalHeight for TreeHistory 0, 5, and 10, except selected problem-code no-height cases.")
Update-HeightProtocolControlState

$heightProtocolNoteLabel = New-Object System.Windows.Forms.Label
$heightProtocolNoteLabel.Text = "Use recorded IDBH units: 5.0 inches = 50, 17.0 inches = 170. Legacy height problem codes 72/74/75 and live TreeClass/TreeStatus codes are honored."
$heightProtocolNoteLabel.Location = New-Object System.Drawing.Point(170, 390)
$heightProtocolNoteLabel.Size = New-Object System.Drawing.Size(940, 22)
$heightProtocolNoteLabel.Anchor = "Top,Left"
$heightProtocolNoteLabel.ForeColor = $script:DadaColorMutedText
$optionsGroup.Controls.Add($heightProtocolNoteLabel)

$heightRequiredMinorPlotsLabel = New-Object System.Windows.Forms.Label
$heightRequiredMinorPlotsLabel.Text = "Minor plots where tree height is required"
$heightRequiredMinorPlotsLabel.Location = New-Object System.Drawing.Point(15, 426)
$heightRequiredMinorPlotsLabel.Size = New-Object System.Drawing.Size(250, 24)
$optionsGroup.Controls.Add($heightRequiredMinorPlotsLabel)

$heightRequiredMinorPlotsBox = New-Object System.Windows.Forms.TextBox
$heightRequiredMinorPlotsBox.Location = New-Object System.Drawing.Point(275, 423)
$heightRequiredMinorPlotsBox.Size = New-Object System.Drawing.Size(835, 24)
$heightRequiredMinorPlotsBox.Anchor = "Top,Left"
$heightRequiredMinorPlotsBox.Text = $script:DefaultHeightRequiredMinorPlots
$optionsGroup.Controls.Add($heightRequiredMinorPlotsBox)
$optionsToolTip.SetToolTip($heightRequiredMinorPlotsBox, "Optional. Enter MinorPlot values where tree TotalHeight is required. Leave blank when height requirements apply to all minor plots.")

$heightRequiredMinorPlotsExampleLabel = New-Object System.Windows.Forms.Label
$heightRequiredMinorPlotsExampleLabel.Text = "Example format: 1, 2, 3. Leave blank to apply required-height checks to all minor plots."
$heightRequiredMinorPlotsExampleLabel.Location = New-Object System.Drawing.Point(275, 451)
$heightRequiredMinorPlotsExampleLabel.Size = New-Object System.Drawing.Size(835, 22)
$heightRequiredMinorPlotsExampleLabel.Anchor = "Top,Left"
$heightRequiredMinorPlotsExampleLabel.ForeColor = $script:DadaColorMutedText
$optionsGroup.Controls.Add($heightRequiredMinorPlotsExampleLabel)

$regenMinorPlotSectionLabel = New-Object System.Windows.Forms.Label
$regenMinorPlotSectionLabel.Text = "Regen Protocol"
$regenMinorPlotSectionLabel.Location = New-Object System.Drawing.Point(15, 486)
$regenMinorPlotSectionLabel.Size = New-Object System.Drawing.Size(280, 22)
$regenMinorPlotSectionLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$regenMinorPlotSectionLabel.ForeColor = $script:DadaColorAccent
$optionsGroup.Controls.Add($regenMinorPlotSectionLabel)

$regenClassHeaderLabel = New-Object System.Windows.Forms.Label
$regenClassHeaderLabel.Text = "Regen class"
$regenClassHeaderLabel.Location = New-Object System.Drawing.Point(15, 540)
$regenClassHeaderLabel.Size = New-Object System.Drawing.Size(240, 20)
$regenClassHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$optionsGroup.Controls.Add($regenClassHeaderLabel)

$regenTimberMinorPlotHeaderLabel = New-Object System.Windows.Forms.Label
$regenTimberMinorPlotHeaderLabel.Text = "Timber minor plots"
$regenTimberMinorPlotHeaderLabel.Location = New-Object System.Drawing.Point(315, 540)
$regenTimberMinorPlotHeaderLabel.Size = New-Object System.Drawing.Size(230, 20)
$regenTimberMinorPlotHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$optionsGroup.Controls.Add($regenTimberMinorPlotHeaderLabel)

$regenWoodlandMinorPlotHeaderLabel = New-Object System.Windows.Forms.Label
$regenWoodlandMinorPlotHeaderLabel.Text = "Woodland minor plots"
$regenWoodlandMinorPlotHeaderLabel.Location = New-Object System.Drawing.Point(650, 540)
$regenWoodlandMinorPlotHeaderLabel.Size = New-Object System.Drawing.Size(240, 20)
$regenWoodlandMinorPlotHeaderLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$optionsGroup.Controls.Add($regenWoodlandMinorPlotHeaderLabel)

$regenTimberSeedlingMinorPlotsLabel = New-Object System.Windows.Forms.Label
$regenTimberSeedlingMinorPlotsLabel.Text = "Seedlings (IDBH 0)"
$regenTimberSeedlingMinorPlotsLabel.Location = New-Object System.Drawing.Point(15, 568)
$regenTimberSeedlingMinorPlotsLabel.Size = New-Object System.Drawing.Size(240, 24)
$optionsGroup.Controls.Add($regenTimberSeedlingMinorPlotsLabel)

$regenTimberSeedlingMinorPlotsBox = New-Object System.Windows.Forms.TextBox
$regenTimberSeedlingMinorPlotsBox.Location = New-Object System.Drawing.Point(315, 565)
$regenTimberSeedlingMinorPlotsBox.Size = New-Object System.Drawing.Size(240, 24)
$regenTimberSeedlingMinorPlotsBox.Text = $script:DefaultRegenTimberSeedlingMinorPlots
$optionsGroup.Controls.Add($regenTimberSeedlingMinorPlotsBox)

$regenWoodlandSeedlingMinorPlotsLabel = New-Object System.Windows.Forms.Label
$regenWoodlandSeedlingMinorPlotsLabel.Text = "Regen woodland seedlings (IDBH 0)"
$regenWoodlandSeedlingMinorPlotsLabel.Location = New-Object System.Drawing.Point(565, 568)
$regenWoodlandSeedlingMinorPlotsLabel.Size = New-Object System.Drawing.Size(280, 24)
$regenWoodlandSeedlingMinorPlotsLabel.Visible = $false
$optionsGroup.Controls.Add($regenWoodlandSeedlingMinorPlotsLabel)

$regenWoodlandSeedlingMinorPlotsBox = New-Object System.Windows.Forms.TextBox
$regenWoodlandSeedlingMinorPlotsBox.Location = New-Object System.Drawing.Point(650, 565)
$regenWoodlandSeedlingMinorPlotsBox.Size = New-Object System.Drawing.Size(240, 24)
$regenWoodlandSeedlingMinorPlotsBox.Text = $script:DefaultRegenWoodlandSeedlingMinorPlots
$optionsGroup.Controls.Add($regenWoodlandSeedlingMinorPlotsBox)

$regenTimberSapling20MinorPlotsLabel = New-Object System.Windows.Forms.Label
$regenTimberSapling20MinorPlotsLabel.Text = "Saplings (IDBH 20)"
$regenTimberSapling20MinorPlotsLabel.Location = New-Object System.Drawing.Point(15, 596)
$regenTimberSapling20MinorPlotsLabel.Size = New-Object System.Drawing.Size(240, 24)
$optionsGroup.Controls.Add($regenTimberSapling20MinorPlotsLabel)

$regenTimberSapling20MinorPlotsBox = New-Object System.Windows.Forms.TextBox
$regenTimberSapling20MinorPlotsBox.Location = New-Object System.Drawing.Point(315, 593)
$regenTimberSapling20MinorPlotsBox.Size = New-Object System.Drawing.Size(240, 24)
$regenTimberSapling20MinorPlotsBox.Text = $script:DefaultRegenTimberSapling20MinorPlots
$optionsGroup.Controls.Add($regenTimberSapling20MinorPlotsBox)

$regenWoodlandSapling20MinorPlotsLabel = New-Object System.Windows.Forms.Label
$regenWoodlandSapling20MinorPlotsLabel.Text = "Regen woodland saplings (IDBH 20)"
$regenWoodlandSapling20MinorPlotsLabel.Location = New-Object System.Drawing.Point(565, 596)
$regenWoodlandSapling20MinorPlotsLabel.Size = New-Object System.Drawing.Size(280, 24)
$regenWoodlandSapling20MinorPlotsLabel.Visible = $false
$optionsGroup.Controls.Add($regenWoodlandSapling20MinorPlotsLabel)

$regenWoodlandSapling20MinorPlotsBox = New-Object System.Windows.Forms.TextBox
$regenWoodlandSapling20MinorPlotsBox.Location = New-Object System.Drawing.Point(650, 593)
$regenWoodlandSapling20MinorPlotsBox.Size = New-Object System.Drawing.Size(240, 24)
$regenWoodlandSapling20MinorPlotsBox.Text = $script:DefaultRegenWoodlandSapling20MinorPlots
$optionsGroup.Controls.Add($regenWoodlandSapling20MinorPlotsBox)

$regenTimberSapling40MinorPlotsLabel = New-Object System.Windows.Forms.Label
$regenTimberSapling40MinorPlotsLabel.Text = "Saplings (IDBH 40)"
$regenTimberSapling40MinorPlotsLabel.Location = New-Object System.Drawing.Point(15, 624)
$regenTimberSapling40MinorPlotsLabel.Size = New-Object System.Drawing.Size(240, 24)
$optionsGroup.Controls.Add($regenTimberSapling40MinorPlotsLabel)

$regenTimberSapling40MinorPlotsBox = New-Object System.Windows.Forms.TextBox
$regenTimberSapling40MinorPlotsBox.Location = New-Object System.Drawing.Point(315, 621)
$regenTimberSapling40MinorPlotsBox.Size = New-Object System.Drawing.Size(240, 24)
$regenTimberSapling40MinorPlotsBox.Text = $script:DefaultRegenTimberSapling40MinorPlots
$optionsGroup.Controls.Add($regenTimberSapling40MinorPlotsBox)

$regenWoodlandSapling40MinorPlotsLabel = New-Object System.Windows.Forms.Label
$regenWoodlandSapling40MinorPlotsLabel.Text = "Regen woodland saplings (IDBH 40)"
$regenWoodlandSapling40MinorPlotsLabel.Location = New-Object System.Drawing.Point(565, 624)
$regenWoodlandSapling40MinorPlotsLabel.Size = New-Object System.Drawing.Size(280, 24)
$regenWoodlandSapling40MinorPlotsLabel.Visible = $false
$optionsGroup.Controls.Add($regenWoodlandSapling40MinorPlotsLabel)

$regenWoodlandSapling40MinorPlotsBox = New-Object System.Windows.Forms.TextBox
$regenWoodlandSapling40MinorPlotsBox.Location = New-Object System.Drawing.Point(650, 621)
$regenWoodlandSapling40MinorPlotsBox.Size = New-Object System.Drawing.Size(240, 24)
$regenWoodlandSapling40MinorPlotsBox.Text = $script:DefaultRegenWoodlandSapling40MinorPlots
$optionsGroup.Controls.Add($regenWoodlandSapling40MinorPlotsBox)

$regenMinorPlotExampleLabel = New-Object System.Windows.Forms.Label
$regenMinorPlotExampleLabel.Text = "Optional regen MinorPlot rules. Example: 1 or 2, 3. Blank categories are not enforced; woodland categories use the woodland species list above."
$regenMinorPlotExampleLabel.Location = New-Object System.Drawing.Point(15, 652)
$regenMinorPlotExampleLabel.Size = New-Object System.Drawing.Size(1095, 22)
$regenMinorPlotExampleLabel.Anchor = "Top,Left"
$regenMinorPlotExampleLabel.ForeColor = $script:DadaColorMutedText
$optionsGroup.Controls.Add($regenMinorPlotExampleLabel)

$optionsToolTip.SetToolTip($regenTimberSeedlingMinorPlotsBox, "Optional. Allowed MinorPlot values for timber regen seedlings where IDBH is 0.")
$optionsToolTip.SetToolTip($regenTimberSapling20MinorPlotsBox, "Optional. Allowed MinorPlot values for timber regen saplings where IDBH is 20.")
$optionsToolTip.SetToolTip($regenTimberSapling40MinorPlotsBox, "Optional. Allowed MinorPlot values for timber regen saplings where IDBH is 40.")
$optionsToolTip.SetToolTip($regenWoodlandSeedlingMinorPlotsBox, "Optional. Allowed MinorPlot values for woodland regen seedlings where IDBH is 0. Uses woodland species codes entered above.")
$optionsToolTip.SetToolTip($regenWoodlandSapling20MinorPlotsBox, "Optional. Allowed MinorPlot values for woodland regen saplings where IDBH is 20. Uses woodland species codes entered above.")
$optionsToolTip.SetToolTip($regenWoodlandSapling40MinorPlotsBox, "Optional. Allowed MinorPlot values for woodland regen saplings where IDBH is 40. Uses woodland species codes entered above.")

$customCodeMetadataNoteLabel = New-Object System.Windows.Forms.Label
$customCodeMetadataNoteLabel.Text = "Only fields marked Active in AppColumns are reviewed for data-entry errors. Excluded fields such as per-acre expansion, CalcSiteIndex, and GMP are skipped even if active."
$customCodeMetadataNoteLabel.Location = New-Object System.Drawing.Point(15, 690)
$customCodeMetadataNoteLabel.Size = New-Object System.Drawing.Size(1095, 38)
$customCodeMetadataNoteLabel.Anchor = "Top,Left"
$optionsGroup.Controls.Add($customCodeMetadataNoteLabel)

function Set-RunOptionControlWidth {
    param(
        [System.Windows.Forms.Control]$Control,
        [int]$MaxWidth,
        [int]$MinimumWidth = 180,
        [int]$RightMargin = 32
    )

    if ($null -eq $Control -or $null -eq $Control.Parent) { return }

    $availableWidth = $Control.Parent.ClientSize.Width - $Control.Left - $RightMargin
    if ($availableWidth -lt $MinimumWidth) { $availableWidth = $MinimumWidth }
    $Control.Width = [Math]::Min($MaxWidth, $availableWidth)
}

function Update-RunOptionControlWidths {
    Set-RunOptionControlWidth -Control $woodlandSpeciesCodesBox -MaxWidth 570 -MinimumWidth 260
    Set-RunOptionControlWidth -Control $woodlandSpeciesExampleLabel -MaxWidth 570 -MinimumWidth 260
    Set-RunOptionControlWidth -Control $heightRareSpeciesCodesBox -MaxWidth 440 -MinimumWidth 220
    Set-RunOptionControlWidth -Control $heightProtocolNoteLabel -MaxWidth 940 -MinimumWidth 420
    Set-RunOptionControlWidth -Control $heightRequiredMinorPlotsBox -MaxWidth 835 -MinimumWidth 320
    Set-RunOptionControlWidth -Control $heightRequiredMinorPlotsExampleLabel -MaxWidth 835 -MinimumWidth 320
    Set-RunOptionControlWidth -Control $regenMinorPlotExampleLabel -MaxWidth 1095 -MinimumWidth 520
    Set-RunOptionControlWidth -Control $customCodeMetadataNoteLabel -MaxWidth 1095 -MinimumWidth 520
}

$optionsGroup.Add_SizeChanged({
    Update-RunOptionControlWidths
})

Update-RunOptionControlWidths

$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Location = New-Object System.Drawing.Point(16, 778)
$actionPanel.Size = New-Object System.Drawing.Size(1192, 470)
$actionPanel.Anchor = "Top,Left"
$actionPanel.BackColor = $script:DadaColorBackground
$mainPanel.Controls.Add($actionPanel)

$previewButton = New-Object System.Windows.Forms.Button
$previewButton.Text = "Preview"
$previewButton.Location = New-Object System.Drawing.Point(20, 22)
$previewButton.Size = New-Object System.Drawing.Size(95, 34)
$previewButton.Visible = $false
$previewButton.Enabled = $false
$actionPanel.Controls.Add($previewButton)

$cleanButton = New-Object System.Windows.Forms.Button
$cleanButton.Text = "Clean table"
$cleanButton.Location = New-Object System.Drawing.Point(125, 22)
$cleanButton.Size = New-Object System.Drawing.Size(105, 34)
$cleanButton.Visible = $false
$cleanButton.Enabled = $false
$actionPanel.Controls.Add($cleanButton)

$workbookChecksButton = New-Object System.Windows.Forms.Button
$workbookChecksButton.Text = "Run"
$workbookChecksButton.Location = New-Object System.Drawing.Point(20, 22)
$workbookChecksButton.Size = New-Object System.Drawing.Size(105, 34)
$actionPanel.Controls.Add($workbookChecksButton)

$clearButton.Location = New-Object System.Drawing.Point(135, 22)
$actionPanel.Controls.Add($clearButton)

$cancelRunButton = New-Object System.Windows.Forms.Button
$cancelRunButton.Text = "Cancel run"
$cancelRunButton.Location = New-Object System.Drawing.Point(245, 22)
$cancelRunButton.Size = New-Object System.Drawing.Size(110, 34)
$cancelRunButton.Visible = $false
$cancelRunButton.Enabled = $false
$actionPanel.Controls.Add($cancelRunButton)

$bitnessLabel = New-Object System.Windows.Forms.Label
$bitnessLabel.Location = New-Object System.Drawing.Point(20, 62)
$bitnessLabel.Size = New-Object System.Drawing.Size(820, 24)
$bitnessLabel.Anchor = "Top,Left"
$bitnessLabel.Text = "Running as " + $(if ([Environment]::Is64BitProcess) { "64-bit" } else { "32-bit" }) + " process"
$bitnessLabel.ForeColor = $script:DadaColorMutedText
$actionPanel.Controls.Add($bitnessLabel)

$progressStatusLabel = New-Object System.Windows.Forms.Label
$progressStatusLabel.Location = New-Object System.Drawing.Point(20, 86)
$progressStatusLabel.Size = New-Object System.Drawing.Size(820, 20)
$progressStatusLabel.Anchor = "Top,Left"
$progressStatusLabel.Text = "Ready"
$progressStatusLabel.ForeColor = $script:DadaColorText
$actionPanel.Controls.Add($progressStatusLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 108)
$progressBar.Size = New-Object System.Drawing.Size(820, 16)
$progressBar.Anchor = "Top,Left"
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$actionPanel.Controls.Add($progressBar)

$aiOptionsGroup = New-Object System.Windows.Forms.GroupBox
$aiOptionsGroup.Text = "Optional AI guidance"
$aiOptionsGroup.Location = New-Object System.Drawing.Point(16, 14)
$aiOptionsGroup.Size = New-Object System.Drawing.Size(820, 245)
$aiOptionsGroup.Anchor = "Top,Left"
$aiOptionsGroup.BackColor = $script:DadaColorAiSurfaceMuted
$aiOptionsGroup.ForeColor = $script:DadaColorAiAccent
$aiPanel.Controls.Add($aiOptionsGroup)

$aiOptionalNoteLabel = New-Object System.Windows.Forms.Label
$aiOptionalNoteLabel.Text = "Optional: leave this off for a normal cleaning run. No AI setup is required. AI can answer chat questions and write AIMessage guidance in the Excel export."
$aiOptionalNoteLabel.Location = New-Object System.Drawing.Point(16, 28)
$aiOptionalNoteLabel.Size = New-Object System.Drawing.Size(760, 34)
$aiOptionalNoteLabel.Anchor = "Top,Left"
$aiOptionalNoteLabel.ForeColor = $script:DadaColorMutedText
$aiOptionsGroup.Controls.Add($aiOptionalNoteLabel)

$aiEnabledCheck = New-Object System.Windows.Forms.CheckBox
$aiEnabledCheck.Text = "AI helper on"
$aiEnabledCheck.Location = New-Object System.Drawing.Point(16, 76)
$aiEnabledCheck.Size = New-Object System.Drawing.Size(115, 24)
$aiEnabledCheck.Checked = $topAiEnabledCheck.Checked
$aiOptionsGroup.Controls.Add($aiEnabledCheck)

$aiHelperButton = New-Object System.Windows.Forms.Button
$aiHelperButton.Text = "Open AI helper"
$aiHelperButton.Location = New-Object System.Drawing.Point(145, 72)
$aiHelperButton.Size = New-Object System.Drawing.Size(130, 30)
$aiOptionsGroup.Controls.Add($aiHelperButton)

$aiExportGuidanceCheck = New-Object System.Windows.Forms.CheckBox
$aiExportGuidanceCheck.Text = "Use AI guidance in export"
$aiExportGuidanceCheck.Location = New-Object System.Drawing.Point(300, 76)
$aiExportGuidanceCheck.Size = New-Object System.Drawing.Size(250, 24)
$aiOptionsGroup.Controls.Add($aiExportGuidanceCheck)
$optionsToolTip.SetToolTip($aiExportGuidanceCheck, "Uses the AI helper, uploaded manual when present, coded cleaning rules, and template codes to add AIMessage guidance to the export.")

$aiCompactExportCheck = New-Object System.Windows.Forms.CheckBox
$aiCompactExportCheck.Text = "Compact AI export"
$aiCompactExportCheck.Location = New-Object System.Drawing.Point(575, 76)
$aiCompactExportCheck.Size = New-Object System.Drawing.Size(180, 24)
$aiCompactExportCheck.Checked = $true
$aiOptionsGroup.Controls.Add($aiCompactExportCheck)
$optionsToolTip.SetToolTip($aiCompactExportCheck, "Sends smaller AI prompts during Run. Leave this on to reduce Azure 400 errors and speed up AI guidance.")

$projectManualLabel.Location = New-Object System.Drawing.Point(16, 118)
$projectManualLabel.Size = New-Object System.Drawing.Size(100, 24)
$aiOptionsGroup.Controls.Add($projectManualLabel)

$projectManualBox.Location = New-Object System.Drawing.Point(120, 115)
$projectManualBox.Size = New-Object System.Drawing.Size(530, 24)
$projectManualBox.Anchor = "Top,Left"
$aiOptionsGroup.Controls.Add($projectManualBox)

$projectManualButton.Location = New-Object System.Drawing.Point(665, 113)
$projectManualButton.Size = New-Object System.Drawing.Size(80, 30)
$projectManualButton.Anchor = "Top,Left"
$aiOptionsGroup.Controls.Add($projectManualButton)

$projectManualAiNoteLabel = New-Object System.Windows.Forms.Label
$projectManualAiNoteLabel.Text = "Manual uploads are used only by AI chat and AIMessage export guidance. They do not change DADA's built-in cleaning checks or decide which records are flagged."
$projectManualAiNoteLabel.Location = New-Object System.Drawing.Point(120, 151)
$projectManualAiNoteLabel.Size = New-Object System.Drawing.Size(640, 42)
$projectManualAiNoteLabel.Anchor = "Top,Left"
$projectManualAiNoteLabel.ForeColor = $script:DadaColorMutedText
$aiOptionsGroup.Controls.Add($projectManualAiNoteLabel)

$optionsToolTip.SetToolTip($projectManualButton, "Optional. Upload a project manual for AI chat and AIMessage export guidance. Built-in cleaning rules are unchanged.")
$optionsToolTip.SetToolTip($projectManualBox, "Manual text is used only as AI context when AI chat or AI export guidance is enabled.")

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(20, 138)
$logBox.Size = New-Object System.Drawing.Size(820, 284)
$logBox.Anchor = "Top,Left,Bottom"
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::White
$logBox.ForeColor = $script:DadaColorText
$logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$actionPanel.Controls.Add($logBox)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = "DisplayedCells"
$grid.Visible = $false

Apply-DadaModernTheme -Control $form
$form.BackColor = $script:DadaColorBackground
$topPanel.BackColor = $script:DadaColorSurface
$mainPanel.BackColor = $script:DadaColorBackground
$aiTabPage.BackColor = $script:DadaColorAiSurface
$aiPanel.BackColor = $script:DadaColorAiSurface
$actionPanel.BackColor = $script:DadaColorBackground
$optionsGroup.BackColor = $script:DadaColorSurface
$aiOptionsGroup.BackColor = $script:DadaColorAiSurfaceMuted
$aiOptionsGroup.ForeColor = $script:DadaColorAiAccent
$headerTitleLabel.ForeColor = $script:DadaColorAccent
$headerSubtitleLabel.ForeColor = $script:DadaColorMutedText
$buildLabel.BackColor = $script:DadaColorSoftAccent
$buildLabel.ForeColor = $script:DadaColorAccent
$creditLabel.ForeColor = $script:DadaColorMutedText
$periodScopeNoteLabel.ForeColor = $script:DadaColorMutedText
$woodlandSpeciesExampleLabel.ForeColor = $script:DadaColorMutedText
$heightProtocolSectionLabel.ForeColor = $script:DadaColorAccent
$heightProtocolNoteLabel.ForeColor = $script:DadaColorMutedText
$heightRequiredMinorPlotsExampleLabel.ForeColor = $script:DadaColorMutedText
$regenMinorPlotSectionLabel.ForeColor = $script:DadaColorAccent
$regenClassHeaderLabel.ForeColor = $script:DadaColorMutedText
$regenTimberMinorPlotHeaderLabel.ForeColor = $script:DadaColorMutedText
$regenWoodlandMinorPlotHeaderLabel.ForeColor = $script:DadaColorMutedText
$regenMinorPlotExampleLabel.ForeColor = $script:DadaColorMutedText
$bitnessLabel.ForeColor = $script:DadaColorMutedText
$aiOptionalNoteLabel.ForeColor = $script:DadaColorAiAccent
$projectManualAiNoteLabel.ForeColor = $script:DadaColorAiAccent
Set-DadaModernButton -Button $browseButton -Style "Secondary"
Set-DadaModernButton -Button $projectManualButton -Style "Secondary"
Set-DadaModernButton -Button $previewButton -Style "Secondary"
Set-DadaModernButton -Button $cleanButton -Style "Secondary"
Set-DadaModernButton -Button $workbookChecksButton -Style "Primary"
Set-DadaModernButton -Button $clearButton -Style "Secondary"
Set-DadaModernButton -Button $cancelRunButton -Style "Danger"
Set-DadaModernButton -Button $topGuideButton -Style "Primary"
Set-DadaModernButton -Button $topAiHelperButton -Style "Secondary"
Set-DadaModernButton -Button $topCancelRunButton -Style "Danger"
Set-DadaModernButton -Button $aiHelperButton -Style "Secondary"

function Connect-SelectedDatabase {
    try {
        $selectedPath = $pathBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($selectedPath)) {
            throw "Choose an Access database first."
        }

        Set-AppProgress -Status "Connecting to database..." -Indeterminate
        $tableCombo.Items.Clear()
        $script:DbPath = $selectedPath
        $script:ConnectionString = $null
        $script:ConnectionString = New-AccessConnectionString -Path $script:DbPath -Password $passwordBox.Text
        Refresh-Tables
        Set-AppProgress -Status "Database connected." -Percent 100
        Add-Log "Connected to $script:DbPath"
        return $true
    }
    catch {
        $script:ConnectionString = $null
        Set-AppProgress -Status "Connection failed." -Percent 0
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Connection failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        Add-Log "Connection failed: $($_.Exception.Message)"
        return $false
    }
}

function Clear-SelectedInputs {
    if (Test-GuideRunActive) {
        [System.Windows.Forms.MessageBox]::Show("Wait for the current run to finish or cancel it before clearing.", $script:AppName) | Out-Null
        return
    }

    $script:DbPath = $null
    $script:ConnectionString = $null
    $script:ColumnsByTable = @{}
    $script:ProjectManualPath = ""
    $script:ProjectManualText = ""
    $script:ProjectManualLoadedFrom = ""
    $script:ProjectManualDigestByKey = @{}

    foreach ($textControl in @($pathBox, $projectManualBox, $passwordBox)) {
        try {
            if ($null -ne $textControl) { $textControl.Clear() }
        }
        catch {
        }
    }

    foreach ($combo in @($tableCombo, $speciesCombo, $dbhCombo, $heightCombo, $plotCombo, $treeCombo)) {
        try {
            if ($null -ne $combo) { $combo.Items.Clear() }
        }
        catch {
        }
    }

    try {
        if ($null -ne $grid) { $grid.DataSource = $null }
    }
    catch {
    }

    try {
        if ($null -ne $woodlandSpeciesCodesBox) { $woodlandSpeciesCodesBox.Text = $script:DefaultWoodlandSpeciesCodes }
    }
    catch {
    }

    try {
        if ($null -ne $heightRequiredMinorPlotsBox) { $heightRequiredMinorPlotsBox.Text = $script:DefaultHeightRequiredMinorPlots }
    }
    catch {
    }

    try {
        if ($null -ne $problem127HeightCheck) { $problem127HeightCheck.Checked = $true }
        if ($null -ne $problem128HeightCheck) { $problem128HeightCheck.Checked = $true }
        if ($null -ne $problem123HeightCheck) { $problem123HeightCheck.Checked = $true }
        if ($null -ne $heightProtocolCombo) { $heightProtocolCombo.SelectedIndex = 0 }
        if ($null -ne $heightSubsampleMinimumBox) { $heightSubsampleMinimumBox.Value = 2 }
        if ($null -ne $heightSubsampleMinIdbhBox) { $heightSubsampleMinIdbhBox.Value = 50 }
        if ($null -ne $heightSubsampleAllAtOrAboveCheck) { $heightSubsampleAllAtOrAboveCheck.Checked = $false }
        if ($null -ne $heightSubsampleAllAtOrAboveBox) { $heightSubsampleAllAtOrAboveBox.Value = 170 }
        if ($null -ne $heightRareSpeciesCodesBox) { $heightRareSpeciesCodesBox.Clear() }
        Update-HeightProtocolControlState
    }
    catch {
    }

    try {
        if ($null -ne $periodScopeCheck) { $periodScopeCheck.Checked = $true }
        if ($null -ne $singleMeasurementProjectCheck) { $singleMeasurementProjectCheck.Checked = $false }
        if ($null -ne $currentPeriodBox) { $currentPeriodBox.Value = 0 }
        if ($null -ne $pastPeriodBox) { $pastPeriodBox.Value = 0 }
        Update-PeriodScopeControlState
    }
    catch {
    }

    foreach ($textControl in @(
        $regenTimberSeedlingMinorPlotsBox,
        $regenTimberSapling20MinorPlotsBox,
        $regenTimberSapling40MinorPlotsBox,
        $regenWoodlandSeedlingMinorPlotsBox,
        $regenWoodlandSapling20MinorPlotsBox,
        $regenWoodlandSapling40MinorPlotsBox
    )) {
        try {
            if ($null -ne $textControl) { $textControl.Clear() }
        }
        catch {
        }
    }

    Set-AppProgress -Status "Cleared selected database, project manual, woodland species codes, and minor-plot rules." -Percent 0
    Add-Log "Cleared selected database, project manual, woodland species codes, and minor-plot rules."
}

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Access databases (*.mdb;*.accdb)|*.mdb;*.accdb|All files (*.*)|*.*"
    $dialog.Title = "Choose an Access database"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $pathBox.Text = $dialog.FileName
        [void](Connect-SelectedDatabase)
    }
    $dialog.Dispose()
})

$projectManualButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Project manual (*.pdf;*.docx;*.xlsx;*.xlsm;*.xls;*.txt;*.md;*.csv;*.json)|*.pdf;*.docx;*.xlsx;*.xlsm;*.xls;*.txt;*.md;*.csv;*.json|All files (*.*)|*.*"
    $dialog.Title = "Choose a project field manual"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            Set-AppProgress -Status "Loading project manual..." -Indeterminate
            Set-ProjectManualFile -Path $dialog.FileName
            Set-AppProgress -Status "Project manual loaded." -Percent 100
            Add-Log "Project manual loaded: $($dialog.FileName)"
        }
        catch {
            Set-AppProgress -Status "Manual upload failed." -Percent 0
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Manual upload failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            Add-Log "Manual upload failed: $($_.Exception.Message)"
        }
    }
    $dialog.Dispose()
})

$clearButton.Add_Click({
    Clear-SelectedInputs
})

$tableCombo.Add_SelectedIndexChanged({
    try {
        Refresh-ColumnsForSelectedTable
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Could not load table", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        Add-Log "Could not load table: $($_.Exception.Message)"
    }
})

$previewButton.Add_Click({
    try {
        Show-Preview
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Preview failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        Add-Log "Preview failed: $($_.Exception.Message)"
    }
})

$cleanButton.Add_Click({
    Clean-SelectedTable
})

$workbookChecksButton.Add_Click({
    Run-CfiWorkbookChecks
})

$topGuideButton.Add_Click({
    Run-CfiWorkbookChecks
})

$cancelRunButton.Add_Click({
    Request-GuideRunCancel
})

$topCancelRunButton.Add_Click({
    Request-GuideRunCancel
})

$aiHelperButton.Add_Click({
    Show-AiHelperDialog
})

$topAiHelperButton.Add_Click({
    Show-AiHelperDialog
})

$aiEnabledCheck.Add_CheckedChanged({
    if ($topAiEnabledCheck.Checked -ne $aiEnabledCheck.Checked) {
        $topAiEnabledCheck.Checked = $aiEnabledCheck.Checked
    }
})

$topAiEnabledCheck.Add_CheckedChanged({
    if ($aiEnabledCheck.Checked -ne $topAiEnabledCheck.Checked) {
        $aiEnabledCheck.Checked = $topAiEnabledCheck.Checked
    }
})

$form.Add_FormClosing({
    param($sender, $eventArgs)
    Confirm-AppCloseWhileRunActive -EventArgs $eventArgs
})

if ([Environment]::Is64BitProcess) {
    Add-Log "Use Run-DADA-DatabaseDad-32bit.bat if this database needs 32-bit Access drivers."
}
else {
    Add-Log "Ready. Running in a 32-bit process."
}

Close-DadaSplashScreen -Splash $script:StartupSplash
$script:StartupSplash = $null

[void]$form.ShowDialog()
