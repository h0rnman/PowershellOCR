Function Get-OCRText() {
[CmdletBinding()]

    Param(
        [parameter(Mandatory=$false, Position=0)]
        [string[]]$Images,
        [parameter(Mandatory=$false)]
        [string]$Path
    )

    $document = New-Object -ComObject MODI.Document
    $temp = New-Object -ComObject MODI.Document
    $output = @()

    if ($Path) {
        foreach ($file in Get-ChildItem $Path) {
            try {
                $temp = New-Object System.Drawing.Bitmap -ArgumentList $file.FullName
                $Images.Add($file.FullName)
            }
            catch {
                Write-Verbose ("{0} is not a valid bitmap file" -f $file.FullName)
            }
        }

    }

    if ($images.Count -eq 0) {return}

    foreach ($image in $images) {

        # Create a bitmap object from the file(s).  MODI can generally only operate on BMP and JPG
        # images, but if we convert the file to a Bitmap object, it can be any supprted Windows type.
        $bitmap = New-Object System.Drawing.Bitmap -ArgumentList $image

        if ($document.images.Count -eq 0) {
            # First image in the array, so add this to the base $document
            $document.Create($image)
        }
        else {
            # Not the first image, so add to ImageAdder so that we have a valid reference
            # This is necessary because there is no Powershell equivalent to the Nothing VBScript keyword
            # and the Add method needs either an Image object or an empty object reference, which Powershell
            # doesn't like.
            $temp.create($image)

            # Add the current object $ImageAdder.Images[0] to the $document stack
            $document.Images.Add($temp.Images[0], $document.Images[0])
        }

    }

    $document.OCR()

    for ([int]$i = $document.Images.Count - 1; $i -ge 0; $i-- ) {
        $tempOut = "" | select Line,Contents
        $tempOut.Line = ($document.Images.Count - $i)
        $tempOut.Contents = $document.Images.Item($i).Layout.Text.Trim()
        $output += $tempOut
    }

    return $output
}