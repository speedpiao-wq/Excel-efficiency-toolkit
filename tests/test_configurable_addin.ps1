$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$addInPath = Join-Path $projectRoot 'outputs\文本框效率工具-可配置版.xlam'
$testId = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$testZip = Join-Path $projectRoot "work\test-$testId.zip"
$testPackage = Join-Path $projectRoot "work\test-$testId"

function Release-ComObject {
    param([Parameter(ValueFromPipeline = $true)] $ComObject)
    if ($null -ne $ComObject) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Assert-Equal {
    param($Actual, $Expected, [string] $Label)
    if ([double] $Actual -ne [double] $Expected) {
        throw "$Label failed. Expected $Expected, got $Actual."
    }
}

function Color-Long {
    param([int] $Red, [int] $Green, [int] $Blue)
    return $Red + (256 * $Green) + (65536 * $Blue)
}

$excel = $null
$addInWorkbook = $null
$testWorkbook = $null
$sheet = $null
$textBox = $null
$formComponent = $null
$designer = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $addInWorkbook = $excel.Workbooks.Open($addInPath)
    $formComponent = $addInWorkbook.VBProject.VBComponents.Item('frmTextBoxSettings')
    $designer = $formComponent.Designer

    $requiredControls = @(
        'txtFontSize',
        'txtFontColor',
        'txtFillColor',
        'chkBold',
        'lblFontPreview',
        'lblFillPreview',
        'btnSave',
        'btnCancel',
        'btnReset'
    )
    foreach ($controlName in $requiredControls) {
        if ($null -eq $designer.Controls.Item($controlName)) {
            throw "Missing settings control: $controlName"
        }
    }

    if (-not [bool] $excel.Run("'文本框效率工具-可配置版.xlam'!SelfTestSettingsForm")) {
        throw 'Settings form initialization self-test failed.'
    }

    $excel.Run("'文本框效率工具-可配置版.xlam'!SaveFormatterSettings", 16, '#C00000', '#FFF2CC', $true)

    $testWorkbook = $excel.Workbooks.Add()
    $sheet = $testWorkbook.Worksheets.Item(1)
    $textBox = $sheet.Shapes.AddTextbox(1, 20, 20, 280, 90)
    $textBox.TextFrame2.TextRange.Text = "可配置版测试`r混合文字格式"
    $textBox.TextFrame2.TextRange.Font.Size = 9
    $textBox.TextFrame2.TextRange.Font.Bold = 0
    $textBox.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 255
    $textBox.Fill.ForeColor.RGB = 65535
    $textBox.Select()

    $excel.Run("'文本框效率工具-可配置版.xlam'!FormatSelectedTextBoxes")

    $expectedFontColor = Color-Long 192 0 0
    $expectedFillColor = Color-Long 255 242 204
    Assert-Equal $textBox.TextFrame2.TextRange.Font.Size 16 'Configured font size'
    Assert-Equal $textBox.TextFrame2.TextRange.Font.Bold -1 'Configured bold'
    Assert-Equal $textBox.TextFrame2.TextRange.Font.Fill.ForeColor.RGB $expectedFontColor 'Configured font color'
    Assert-Equal $textBox.TextFrame.Characters().Font.Color $expectedFontColor 'Legacy font color'
    Assert-Equal $textBox.Fill.ForeColor.RGB $expectedFillColor 'Configured fill color'
    Assert-Equal $textBox.Fill.Transparency 0 'Configured fill transparency'

    $sheet.Range('A1').Select()
    $excel.Run("'文本框效率工具-可配置版.xlam'!FormatSelectedTextBoxes")

    $excel.Run("'文本框效率工具-可配置版.xlam'!ResetFormatterSettings")
    Assert-Equal ($excel.Run("'文本框效率工具-可配置版.xlam'!GetConfiguredFontSize")) 14 'Reset font size'

    Copy-Item -LiteralPath $addInPath -Destination $testZip
    New-Item -ItemType Directory -Path $testPackage | Out-Null
    Expand-Archive -LiteralPath $testZip -DestinationPath $testPackage
    [xml] $relations = Get-Content -LiteralPath (Join-Path $testPackage '_rels\.rels') -Raw
    [xml] $ribbon = Get-Content -LiteralPath (Join-Path $testPackage 'customUI\customUI.xml') -Raw
    $ribbonRelationship = $relations.Relationships.Relationship | Where-Object { $_.Type -like '*ui/extensibility' }
    if ($ribbonRelationship.Target -ne 'customUI/customUI.xml') {
        throw 'Ribbon relationship is missing or invalid.'
    }
    $buttons = @($ribbon.customUI.ribbon.tabs.tab.group.button)
    if ($buttons.Count -ne 2) {
        throw "Expected 2 Ribbon buttons, got $($buttons.Count)."
    }

    [pscustomobject]@{
        AddInOpened = $true
        SettingsFormControls = $requiredControls.Count
        SettingsFormInitialized = $true
        ConfiguredFontSize = $textBox.TextFrame2.TextRange.Font.Size
        ConfiguredFontColor = ('#{0:X6}' -f 0xC00000)
        ConfiguredFillColor = ('#{0:X6}' -f 0xFFF2CC)
        CellSelectionNoError = $true
        RibbonButtons = ($buttons.label -join ', ')
        DefaultsRestored = $true
    }
}
finally {
    if ($null -ne $testWorkbook) {
        $testWorkbook.Close($false)
    }
    if ($null -ne $addInWorkbook) {
        $addInWorkbook.Close($false)
    }
    if ($null -ne $excel) {
        $excel.Quit()
    }
    $designer | Release-ComObject
    $formComponent | Release-ComObject
    $textBox | Release-ComObject
    $sheet | Release-ComObject
    $testWorkbook | Release-ComObject
    $addInWorkbook | Release-ComObject
    $excel | Release-ComObject
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
