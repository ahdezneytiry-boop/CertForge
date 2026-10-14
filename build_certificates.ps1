$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataPath = Join-Path $root "certificate_data.json"
$bgPath = Join-Path $root "extracted_media\image1.png"
$outDir = Join-Path $root "outputs"
$docxPath = Join-Path $outDir "校级奖状_16张.docx"
$pdfPath = Join-Path $outDir "校级奖状_16张.pdf"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$payload = Get-Content -Path $dataPath -Encoding UTF8 -Raw | ConvertFrom-Json
$records = @($payload.records)

function ColorValue([string]$hex) {
    $h = $hex.TrimStart("#")
    $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
    return ($r + ($g * 256) + ($b * 65536))
}

function WorkTitleSize([string]$text) {
    $len = $text.Length
    if ($len -le 28) { return 23 }
    if ($len -le 42) { return 20 }
    if ($len -le 58) { return 18 }
    return 16.5
}

$wdPageBreak = 7
$wdStory = 6
$wdWrapBehind = 5
$wdWrapNone = 3
$wdRelativePage = 1
$msoTextOrientationHorizontal = 1
$msoFalse = 0
$msoTrue = -1
$msoAnchorMiddle = 3
$msoAnchorTop = 1
$wdAlignLeft = 0
$wdAlignCenter = 1
$wdFormatXMLDocument = 16
$wdExportFormatPDF = 17

$pageW = 841.89
$pageH = 595.28

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()
    $doc.PageSetup.PageWidth = $pageW
    $doc.PageSetup.PageHeight = $pageH
    $doc.PageSetup.TopMargin = 0
    $doc.PageSetup.BottomMargin = 0
    $doc.PageSetup.LeftMargin = 0
    $doc.PageSetup.RightMargin = 0
    $doc.PageSetup.HeaderDistance = 0
    $doc.PageSetup.FooterDistance = 0
    $doc.Content.Font.NameFarEast = "Microsoft YaHei"
    $doc.Content.ParagraphFormat.SpaceBefore = 0
    $doc.Content.ParagraphFormat.SpaceAfter = 0

    function Add-Textbox {
        param(
            [string]$Text,
            [double]$Left,
            [double]$Top,
            [double]$Width,
            [double]$Height,
            [double]$Size,
            [string]$Color,
            [bool]$Bold = $false,
            [int]$Align = $wdAlignCenter,
            [string]$Font = "Microsoft YaHei",
            [int]$Vertical = $msoAnchorMiddle
        )

        $anchor = $word.Selection.Range
        $shape = $doc.Shapes.AddTextbox($msoTextOrientationHorizontal, $Left, $Top, $Width, $Height, $anchor)
        $shape.RelativeHorizontalPosition = $wdRelativePage
        $shape.RelativeVerticalPosition = $wdRelativePage
        $shape.Left = $Left
        $shape.Top = $Top
        $shape.Width = $Width
        $shape.Height = $Height
        $shape.WrapFormat.Type = $wdWrapNone
        $shape.Line.Visible = $msoFalse
        $shape.Fill.Visible = $msoFalse
        $shape.TextFrame.MarginLeft = 0
        $shape.TextFrame.MarginRight = 0
        $shape.TextFrame.MarginTop = 0
        $shape.TextFrame.MarginBottom = 0
        $shape.TextFrame.VerticalAnchor = $Vertical
        $shape.TextFrame.WordWrap = $msoTrue

        $range = $shape.TextFrame.TextRange
        $range.Text = $Text
        $range.Font.Name = $Font
        $range.Font.NameFarEast = $Font
        $range.Font.Size = $Size
        $range.Font.Bold = $(if ($Bold) { 1 } else { 0 })
        $range.Font.Color = ColorValue $Color
        $range.ParagraphFormat.Alignment = $Align
        $range.ParagraphFormat.SpaceBefore = 0
        $range.ParagraphFormat.SpaceAfter = 0
        return $shape
    }

    function Add-Background {
        $anchor = $word.Selection.Range
        $pic = $doc.Shapes.AddPicture($bgPath, $false, $true, 0, 0, $pageW, $pageH, $anchor)
        $pic.RelativeHorizontalPosition = $wdRelativePage
        $pic.RelativeVerticalPosition = $wdRelativePage
        $pic.Left = 0
        $pic.Top = 0
        $pic.Width = $pageW
        $pic.Height = $pageH
        $pic.WrapFormat.Type = $wdWrapBehind
        $pic.ZOrder(5) | Out-Null
    }

    for ($i = 0; $i -lt $records.Count; $i++) {
        if ($i -gt 0) {
            $word.Selection.EndKey($wdStory) | Out-Null
            $word.Selection.InsertBreak($wdPageBreak)
        }

        $r = $records[$i]
        Add-Background

        Add-Textbox -Text $payload.event -Left 155 -Top 57 -Width 532 -Height 24 -Size 13.5 -Color "#9E6A19" -Bold $false | Out-Null
        Add-Textbox -Text "荣誉证书" -Left 0 -Top 94 -Width $pageW -Height 64 -Size 45 -Color "#8A2D1F" -Bold $true -Font "SimHei" | Out-Null

        $teamLine = "$($payload.issuer)参赛团队「$($r.team_name)」的"
        Add-Textbox -Text $teamLine -Left 110 -Top 184 -Width 622 -Height 30 -Size 19 -Color "#1D1D1D" -Bold $true | Out-Null

        $titleText = "《$($r.work_title)》"
        $titleSize = WorkTitleSize $titleText
        Add-Textbox -Text $titleText -Left 108 -Top 220 -Width 626 -Height 74 -Size $titleSize -Color "#111111" -Bold $true -Vertical $msoAnchorMiddle | Out-Null

        $sentence = "在 $($payload.event) 中，荣获"
        Add-Textbox -Text $sentence -Left 120 -Top 303 -Width 600 -Height 32 -Size 18 -Color "#222222" | Out-Null

        Add-Textbox -Text $r.award_display -Left 0 -Top 331 -Width $pageW -Height 42 -Size 24 -Color "#8A2D1F" -Bold $true -Font "SimHei" | Out-Null
        Add-Textbox -Text "特发此证，以资鼓励。" -Left 0 -Top 380 -Width $pageW -Height 28 -Size 17 -Color "#222222" | Out-Null

        Add-Textbox -Text "团队成员：$($r.members)" -Left 158 -Top 427 -Width 526 -Height 28 -Size 15.5 -Color "#1A1A1A" -Align $wdAlignLeft -Vertical $msoAnchorMiddle | Out-Null
        Add-Textbox -Text "指导教师：$($r.teacher)" -Left 158 -Top 459 -Width 526 -Height 28 -Size 15.5 -Color "#1A1A1A" -Align $wdAlignLeft -Vertical $msoAnchorMiddle | Out-Null

        Add-Textbox -Text "颁发单位：$($payload.issuer)" -Left 470 -Top 491 -Width 250 -Height 24 -Size 14.5 -Color "#1A1A1A" -Align $wdAlignLeft | Out-Null
        Add-Textbox -Text "颁发日期：$($payload.date)" -Left 470 -Top 519 -Width 250 -Height 24 -Size 14.5 -Color "#1A1A1A" -Align $wdAlignLeft | Out-Null
    }

    if (Test-Path -Path $docxPath) { Remove-Item -Path $docxPath -Force }
    if (Test-Path -Path $pdfPath) { Remove-Item -Path $pdfPath -Force }

    $doc.SaveAs2($docxPath, $wdFormatXMLDocument)
    $doc.ExportAsFixedFormat($pdfPath, $wdExportFormatPDF)
    $doc.Close($false)
    $word.Quit()

    Write-Output "Saved DOCX: $docxPath"
    Write-Output "Saved PDF: $pdfPath"
} finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) | Out-Null } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($word -ne $null) {
        try { $word.Quit() | Out-Null } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
