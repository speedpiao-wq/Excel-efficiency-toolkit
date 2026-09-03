$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$generatedPath = Join-Path $projectRoot 'work\skill-forward-test\Skill生成测试.xlam'
$mainPath = Join-Path $projectRoot 'outputs\文本框效率工具-可配置版.xlam'
$builderPath = Join-Path $projectRoot 'skill\excel-global-addin-maker\scripts\build_textbox_formatter.ps1'

& $builderPath `
    -OutputPath $generatedPath `
    -DefaultFontSize 18 `
    -DefaultFontHex '#7030A0' `
    -DefaultFillHex '#E2F0D9' `
    -DefaultBold $false

function Color-Long {
    param([int] $Red, [int] $Green, [int] $Blue)
    return $Red + (256 * $Green) + (65536 * $Blue)
}

function Assert-Equal {
    param($Actual, $Expected, [string] $Label)
    if ([double] $Actual -ne [double] $Expected) {
        throw "$Label failed. Expected $Expected, got $Actual."
    }
}

$excel = $null
$generated = $null
$main = $null
$testWorkbook = $null
$sheet = $null
$textBox = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $generated = $excel.Workbooks.Open($generatedPath)

    $excel.Run("'Skill生成测试.xlam'!ResetFormatterSettings")
    Assert-Equal ($excel.Run("'Skill生成测试.xlam'!GetConfiguredFontSize")) 18 'Generated default font size'
    if ($excel.Run("'Skill生成测试.xlam'!GetConfiguredFontHex") -ne '#7030A0') {
        throw 'Generated default font color failed.'
    }
    if ($excel.Run("'Skill生成测试.xlam'!GetConfiguredFillHex") -ne '#E2F0D9') {
        throw 'Generated default fill color failed.'
    }
    if ([bool] $excel.Run("'Skill生成测试.xlam'!GetConfiguredBold")) {
        throw 'Generated default bold setting failed.'
    }

    $testWorkbook = $excel.Workbooks.Add()
    $sheet = $testWorkbook.Worksheets.Item(1)
    $textBox = $sheet.Shapes.AddTextbox(1, 10, 10, 220, 60)
    $textBox.TextFrame2.TextRange.Text = 'Skill forward test'
    $textBox.Select()
    $excel.Run("'Skill生成测试.xlam'!FormatSelectedTextBoxes")

    Assert-Equal $textBox.TextFrame2.TextRange.Font.Size 18 'Generated add-in font size'
    Assert-Equal $textBox.TextFrame2.TextRange.Font.Bold 0 'Generated add-in bold off'
    Assert-Equal $textBox.TextFrame2.TextRange.Font.Fill.ForeColor.RGB (Color-Long 112 48 160) 'Generated add-in font color'
    Assert-Equal $textBox.Fill.ForeColor.RGB (Color-Long 226 240 217) 'Generated add-in fill color'

    [pscustomobject]@{
        SkillBuilder = 'Passed'
        FontSize = 18
        FontColor = '#7030A0'
        FillColor = '#E2F0D9'
        Bold = $false
    }
}
finally {
    if ($null -ne $testWorkbook) { $testWorkbook.Close($false) }
    if ($null -ne $generated) { $generated.Close($false) }

    # Restore the main deliverable's defaults after the isolated forward test.
    if ($null -ne $excel) {
        $main = $excel.Workbooks.Open($mainPath)
        $excel.Run("'文本框效率工具-可配置版.xlam'!ResetFormatterSettings")
        $main.Close($false)
        $excel.Quit()
    }
}
