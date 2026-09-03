param(
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $projectRoot 'outputs\文本框效率工具-可配置版.xlam'
}
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
$modulePath = Join-Path $projectRoot 'src\vba\modTextBoxFormatter.bas'
$formCodePath = Join-Path $projectRoot 'src\vba\frmTextBoxSettings-code.txt'
$ribbonPath = Join-Path $projectRoot 'src\ribbon\customUI.xml'
$workRoot = Join-Path $projectRoot 'work'
$buildId = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$packageFolder = Join-Path $workRoot "package-$buildId"
$packageZip = Join-Path $workRoot "package-$buildId.zip"

function Release-ComObject {
    param([Parameter(ValueFromPipeline = $true)] $ComObject)
    if ($null -ne $ComObject) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Add-FormControl {
    param(
        [Parameter(Mandatory = $true)] $Designer,
        [Parameter(Mandatory = $true)][string] $ProgId,
        [Parameter(Mandatory = $true)][string] $Name,
        [Parameter(Mandatory = $true)][single] $Left,
        [Parameter(Mandatory = $true)][single] $Top,
        [Parameter(Mandatory = $true)][single] $Width,
        [Parameter(Mandatory = $true)][single] $Height,
        [string] $Caption = $null
    )

    $control = $Designer.Controls.Add($ProgId, $Name, $true)
    $control.Left = $Left
    $control.Top = $Top
    $control.Width = $Width
    $control.Height = $Height
    if ($PSBoundParameters.ContainsKey('Caption')) {
        $control.Caption = $Caption
    }
    # The control stays owned by the form designer; callers retrieve it by name.
}

function Add-RibbonCustomization {
    param(
        [Parameter(Mandatory = $true)][string] $AddInPath,
        [Parameter(Mandatory = $true)][string] $RibbonXmlPath,
        [Parameter(Mandatory = $true)][string] $TemporaryZipPath,
        [Parameter(Mandatory = $true)][string] $TemporaryFolder
    )

    Copy-Item -LiteralPath $AddInPath -Destination $TemporaryZipPath
    New-Item -ItemType Directory -Path $TemporaryFolder | Out-Null
    Expand-Archive -LiteralPath $TemporaryZipPath -DestinationPath $TemporaryFolder

    $customUiFolder = Join-Path $TemporaryFolder 'customUI'
    New-Item -ItemType Directory -Path $customUiFolder | Out-Null
    Copy-Item -LiteralPath $RibbonXmlPath -Destination (Join-Path $customUiFolder 'customUI.xml')

    $relationsPath = Join-Path $TemporaryFolder '_rels\.rels'
    [xml] $relationsXml = Get-Content -LiteralPath $relationsPath -Raw
    $relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $existing = $relationsXml.Relationships.Relationship | Where-Object {
        $_.Type -eq 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
    }

    if ($null -eq $existing) {
        $relationship = $relationsXml.CreateElement('Relationship', $relationshipNamespace)
        [void] $relationship.SetAttribute('Id', 'customUIRel')
        [void] $relationship.SetAttribute('Type', 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility')
        [void] $relationship.SetAttribute('Target', 'customUI/customUI.xml')
        [void] $relationsXml.DocumentElement.AppendChild($relationship)
    }

    $xmlSettings = [Xml.XmlWriterSettings]::new()
    $xmlSettings.Encoding = [Text.UTF8Encoding]::new($false)
    $xmlSettings.Indent = $false
    $writer = [Xml.XmlWriter]::Create($relationsPath, $xmlSettings)
    try {
        $relationsXml.Save($writer)
    }
    finally {
        $writer.Dispose()
    }

    Remove-Item -LiteralPath $AddInPath
    Compress-Archive -Path (Join-Path $TemporaryFolder '*') -DestinationPath $TemporaryZipPath -Force
    Move-Item -LiteralPath $TemporaryZipPath -Destination $AddInPath
}

$outputDirectory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
if (-not (Test-Path -LiteralPath $workRoot)) {
    New-Item -ItemType Directory -Path $workRoot | Out-Null
}
if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath
}

$excel = $null
$workbook = $null
$moduleComponent = $null
$formComponent = $null
$designer = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $moduleComponent = $workbook.VBProject.VBComponents.Import($modulePath)
    $moduleComponent.Name = 'modTextBoxFormatter'

    $formComponent = $workbook.VBProject.VBComponents.Add(3)
    $formComponent.Name = 'frmTextBoxSettings'
    $formComponent.Properties.Item('Caption').Value = '文本框格式设置'
    $formComponent.Properties.Item('Width').Value = 360
    $formComponent.Properties.Item('Height').Value = 300
    $formComponent.Properties.Item('BackColor').Value = 16777215

    $designer = $formComponent.Designer

    Add-FormControl $designer 'Forms.Label.1' 'lblTitle' 20 14 310 28 '文本框一键格式'
    $title = $designer.Controls.Item('lblTitle')
    $title.ForeColor = 13382400
    $title.BackStyle = 0

    Add-FormControl $designer 'Forms.Label.1' 'lblIntro' 20 45 310 28 '设置会自动保存，以后打开其他工作簿也会继续使用。'
    $intro = $designer.Controls.Item('lblIntro')
    $intro.ForeColor = 6316128
    $intro.BackStyle = 0

    Add-FormControl $designer 'Forms.Label.1' 'lblFontSize' 24 84 95 20 '字号'
    $fontSizeLabel = $designer.Controls.Item('lblFontSize')
    $fontSizeLabel.BackStyle = 0
    Add-FormControl $designer 'Forms.TextBox.1' 'txtFontSize' 125 80 90 24
    $fontSize = $designer.Controls.Item('txtFontSize')
    $fontSize.TabIndex = 0
    Add-FormControl $designer 'Forms.Label.1' 'lblFontSizeUnit' 225 84 65 20 '例如：14'
    $fontSizeUnit = $designer.Controls.Item('lblFontSizeUnit')
    $fontSizeUnit.ForeColor = 8421504
    $fontSizeUnit.BackStyle = 0

    Add-FormControl $designer 'Forms.Label.1' 'lblFontColor' 24 120 95 20 '字体颜色'
    $fontColorLabel = $designer.Controls.Item('lblFontColor')
    $fontColorLabel.BackStyle = 0
    Add-FormControl $designer 'Forms.TextBox.1' 'txtFontColor' 125 116 90 24
    $fontColor = $designer.Controls.Item('txtFontColor')
    $fontColor.TabIndex = 1
    Add-FormControl $designer 'Forms.Label.1' 'lblFontPreview' 225 117 52 22 ''
    $fontPreview = $designer.Controls.Item('lblFontPreview')
    $fontPreview.BorderStyle = 1

    Add-FormControl $designer 'Forms.Label.1' 'lblFillColor' 24 156 95 20 '填充颜色'
    $fillColorLabel = $designer.Controls.Item('lblFillColor')
    $fillColorLabel.BackStyle = 0
    Add-FormControl $designer 'Forms.TextBox.1' 'txtFillColor' 125 152 90 24
    $fillColor = $designer.Controls.Item('txtFillColor')
    $fillColor.TabIndex = 2
    Add-FormControl $designer 'Forms.Label.1' 'lblFillPreview' 225 153 52 22 ''
    $fillPreview = $designer.Controls.Item('lblFillPreview')
    $fillPreview.BorderStyle = 1

    Add-FormControl $designer 'Forms.CheckBox.1' 'chkBold' 125 187 150 22 '字体加粗'
    $bold = $designer.Controls.Item('chkBold')
    $bold.TabIndex = 3
    $bold.BackStyle = 0

    Add-FormControl $designer 'Forms.Label.1' 'lblTip' 24 216 310 18 '颜色请输入 #RRGGBB，例如 #0033CC。'
    $tip = $designer.Controls.Item('lblTip')
    $tip.ForeColor = 8421504
    $tip.BackStyle = 0

    Add-FormControl $designer 'Forms.CommandButton.1' 'btnReset' 24 246 92 28 '恢复默认'
    $reset = $designer.Controls.Item('btnReset')
    $reset.TabIndex = 6
    Add-FormControl $designer 'Forms.CommandButton.1' 'btnCancel' 150 246 80 28 '取消'
    $cancel = $designer.Controls.Item('btnCancel')
    $cancel.TabIndex = 5
    Add-FormControl $designer 'Forms.CommandButton.1' 'btnSave' 240 246 92 28 '保存设置'
    $save = $designer.Controls.Item('btnSave')
    $save.TabIndex = 4
    $save.Default = $true

    $formCode = Get-Content -LiteralPath $formCodePath -Raw
    $formComponent.CodeModule.AddFromString($formCode)

    $workbook.Title = '文本框效率工具 - 可配置版'
    $workbook.Subject = '全局统一所选文本框格式，并提供免代码设置窗口'
    $workbook.Author = 'Codex'
    $workbook.IsAddin = $true
    $workbook.SaveAs($OutputPath, 55)
    $workbook.Close($false)
    $workbook = $null
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    if ($null -ne $excel) {
        $excel.Quit()
    }
    $designer | Release-ComObject
    $formComponent | Release-ComObject
    $moduleComponent | Release-ComObject
    $workbook | Release-ComObject
    $excel | Release-ComObject
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Add-RibbonCustomization -AddInPath $OutputPath `
                        -RibbonXmlPath $ribbonPath `
                        -TemporaryZipPath $packageZip `
                        -TemporaryFolder $packageFolder

$file = Get-Item -LiteralPath $OutputPath
$hash = Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256
[pscustomobject]@{
    FullName = $file.FullName
    Length = $file.Length
    SHA256 = $hash.Hash
}
