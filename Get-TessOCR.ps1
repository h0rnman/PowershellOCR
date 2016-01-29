Function Get-TessOCR() {
[CmdletBinding()]

Param(
    [parameter(Mandatory=$false,Position=0)]
    [string[]]$Images,
    [parameter(Mandatory=$false)]
    [string]$Path
)

if ($Path) {
    foreach ($file in Get-ChildItem $Path) {
        try {
            $temp = New-Object System.Drawing.Bitmap -ArgumentList $file.fullName
            $Images.Add($temp)
        }
        catch {
            Write-Verbose ("{0} is not a valid image file" -f $file.fullName)
        }
    }
}

# This is the location of your Tesseract binaries and language data
$TesseractLocation = "c:\temp\tesseract"
$results = @()

# The Tesseract.dll file is a .Net wrapper around libtesseract and liblept
Add-Type -Path "$TesseractLocation\lib\Tesseract.dll"

# Create the OCR engine object
$engine = New-Object Tesseract.TesseractEngine("$TesseractLocation\lib\tessdata","eng", [Tesseract.EngineMode]::TesseractAndCube, $null)

foreach ($image in $Images) {
    # Tesseract needs image data in Bitmap format
    $tessImage = New-Object System.Drawing.Bitmap($image)
    $pix = [Tesseract.PixConverter]::ToPix($tessImage)
    $doc = $engine.Process($pix)

    $data = "" | select Text,Confidence,EngineVersion

    # Contrary to what it might seem, this line is where the OCR actually happens
    $data.Text = $doc.GetText()
    $data.Confidence = $doc.GetMeanConfidence()
    $data.EngineVersion = $doc.Engine.Version

    $results += $data

    $doc.Dispose()
    $tessImage.Dispose()

}

$engine.Dispose()
return $results

}