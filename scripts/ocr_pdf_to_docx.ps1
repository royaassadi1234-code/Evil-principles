param(
    [Parameter(Mandatory = $true)]
    [string]$PdfPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDocx,

    [string]$OutputText = "",

    [int]$StartPage = 1,

    [int]$EndPage = 0
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
[void][Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
[void][Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
[void][Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]

$script:AsTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq "AsTask" -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 } |
    Select-Object -First 1

$script:AsTaskAction = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq "AsTask" -and -not $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1 } |
    Select-Object -First 1

function Await-AsyncOperation {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Operation,

        [Parameter(Mandatory = $true)]
        [Type]$ResultType
    )

    $genericMethod = $script:AsTaskGeneric.MakeGenericMethod($ResultType)
    $task = $genericMethod.Invoke($null, @($Operation))
    return $task.GetAwaiter().GetResult()
}

function Await-AsyncAction {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Action
    )

    $task = $script:AsTaskAction.Invoke($null, @($Action))
    $task.GetAwaiter().GetResult() | Out-Null
}

function Escape-Xml {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    return [System.Security.SecurityElement]::Escape($Text)
}

function New-ParagraphXml {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text
    )

    if ($Text.Length -eq 0) {
        return "<w:p/>"
    }

    $escaped = Escape-Xml -Text $Text
    return "<w:p><w:r><w:t xml:space=`"preserve`">$escaped</w:t></w:r></w:p>"
}

function New-PageBreakXml {
    return "<w:p><w:r><w:br w:type=`"page`"/></w:r></w:p>"
}

function Write-ZipEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory = $true)]
        [string]$EntryPath,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $entry = $Archive.CreateEntry($EntryPath)
    $stream = $entry.Open()
    try {
        $writer = New-Object System.IO.StreamWriter($stream, [System.Text.UTF8Encoding]::new($false))
        try {
            $writer.Write($Content)
        }
        finally {
            $writer.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

if (-not $OutputText) {
    $OutputText = [System.IO.Path]::ChangeExtension($OutputDocx, ".txt")
}

$pdfFullPath = [System.IO.Path]::GetFullPath($PdfPath)
$docxFullPath = [System.IO.Path]::GetFullPath($OutputDocx)
$textFullPath = [System.IO.Path]::GetFullPath($OutputText)

$storageFile = Await-AsyncOperation -Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfFullPath)) -ResultType ([Windows.Storage.StorageFile])
$pdf = Await-AsyncOperation -Operation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) -ResultType ([Windows.Data.Pdf.PdfDocument])

if ($pdf.PageCount -lt 1) {
    throw "The PDF has no pages."
}

if ($StartPage -lt 1 -or $StartPage -gt $pdf.PageCount) {
    throw "StartPage must be between 1 and $($pdf.PageCount)."
}

if ($EndPage -eq 0) {
    $EndPage = $pdf.PageCount
}

if ($EndPage -lt $StartPage -or $EndPage -gt $pdf.PageCount) {
    throw "EndPage must be between $StartPage and $($pdf.PageCount)."
}

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new("en-US"))
if (-not $engine) {
    throw "Windows OCR engine could not be created for en-US."
}

$allPages = New-Object System.Collections.Generic.List[string]

for ($pageIndex = ($StartPage - 1); $pageIndex -lt $EndPage; $pageIndex++) {
    $page = $pdf.GetPage($pageIndex)
    try {
        $renderOptions = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
        $targetWidth = [math]::Round([math]::Min([math]::Max($page.Size.Width * 2.5, 1400), 2400))
        $renderOptions.DestinationWidth = [uint32]$targetWidth

        $stream = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
        try {
            Await-AsyncAction -Action ($page.RenderToStreamAsync($stream, $renderOptions))
            $stream.Seek(0) | Out-Null

            $decoder = Await-AsyncOperation -Operation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) -ResultType ([Windows.Graphics.Imaging.BitmapDecoder])
            $bitmap = Await-AsyncOperation -Operation ($decoder.GetSoftwareBitmapAsync()) -ResultType ([Windows.Graphics.Imaging.SoftwareBitmap])
            try {
                $ocrResult = Await-AsyncOperation -Operation ($engine.RecognizeAsync($bitmap)) -ResultType ([Windows.Media.Ocr.OcrResult])
                $pageText = [string]$ocrResult.Text
                if ($null -eq $pageText) {
                    $pageText = ""
                }
                $pageText = $pageText.Trim()
            }
            finally {
                $bitmap.Dispose()
            }
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $page.Dispose()
    }

    $allPages.Add($pageText)
    Write-Host ("OCR page {0}/{1}" -f ($pageIndex + 1), $EndPage)
}

$textBuilder = New-Object System.Text.StringBuilder
$bodyBuilder = New-Object System.Text.StringBuilder

for ($index = 0; $index -lt $allPages.Count; $index++) {
    $pageText = $allPages[$index]
    $lines = $pageText -split "\r?\n"

    foreach ($line in $lines) {
        [void]$textBuilder.AppendLine($line)
        [void]$bodyBuilder.AppendLine((New-ParagraphXml -Text $line))
    }

    if ($index -lt ($allPages.Count - 1)) {
        [void]$textBuilder.AppendLine("")
        [void]$textBuilder.AppendLine("----- PAGE BREAK -----")
        [void]$textBuilder.AppendLine("")
        [void]$bodyBuilder.AppendLine((New-PageBreakXml))
    }
}

$textDir = Split-Path -Parent $textFullPath
$docxDir = Split-Path -Parent $docxFullPath
if ($textDir) {
    [System.IO.Directory]::CreateDirectory($textDir) | Out-Null
}
if ($docxDir) {
    [System.IO.Directory]::CreateDirectory($docxDir) | Out-Null
}

[System.IO.File]::WriteAllText($textFullPath, $textBuilder.ToString(), [System.Text.UTF8Encoding]::new($false))

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    mc:Ignorable="w14 w15 wp14">
  <w:body>
$($bodyBuilder.ToString())
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

$contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

$relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

if (Test-Path -LiteralPath $docxFullPath) {
    Remove-Item -LiteralPath $docxFullPath -Force
}

$zip = [System.IO.Compression.ZipFile]::Open($docxFullPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    Write-ZipEntry -Archive $zip -EntryPath "[Content_Types].xml" -Content $contentTypesXml
    Write-ZipEntry -Archive $zip -EntryPath "_rels/.rels" -Content $relsXml
    Write-ZipEntry -Archive $zip -EntryPath "word/document.xml" -Content $documentXml
}
finally {
    $zip.Dispose()
}

Write-Host ""
Write-Host ("Saved OCR text to: {0}" -f $textFullPath)
Write-Host ("Saved Word file to: {0}" -f $docxFullPath)
