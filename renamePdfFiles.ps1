# install package provider if not available
#Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser

# install itext7
#Install-Module -Name IText7Module -Scope CurrentUser

# get version number with:
#Get-Module -ListAvailable -Name iText7Module

# update module
# Update-Module -Name IText7Module

Add-Type -Path "C:\Users\skromosoeto\Documents\WindowsPowerShell\Modules\IText7Module\1.0.34\itext.kernel.dll"

$folderName = "C:\Users\skromosoeto\Downloads\renamedFieldManagementFiles"

Get-ChildItem -Path $folderName -Filter *.pdf | ForEach-Object {
    $oldFileName = $_.Name
    $fullFileName = $_.FullName
    $pdfReader = [iText.Kernel.Pdf.PdfReader]::new($fullFileName)
    $pdfDocument = [iText.Kernel.Pdf.PdfDocument]::new($pdfReader)
    
    $text = ""
    For ($i = 1; $i -le $pdfDocument.GetNumberOfPages(); $i++) {
        $page = $pdfDocument.GetPage($i)
        $strategy = [iText.Kernel.Pdf.Canvas.Parser.Listener.LocationTextExtractionStrategy]::new()
        $text += [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($page, $strategy)
    }
    
    $pdfDocument.Close()
    
    $reviewCode = $text.Substring(($text.IndexOf('Review Code:') + 12),40)

    if ($reviewCode.Contains("Jaarlijkse")) {
        $newReviewCode = "JaarlijkseEvaluatie"
        $firstName = $text.Substring(($text.IndexOf('First name:') + 11),(($text.IndexOf('Start date:',$text.IndexOf('Start date:')+1) - $text.IndexOf('First name:'))-11))
        $LastName = $text.Substring(($text.IndexOf('Last name:') + 10),(($text.IndexOf('Department:')+1) - $text.IndexOf('Last name:'))-11)
        $endDate = $text.Substring(($text.IndexOf('End Date:') + 10),10)
        $newFilename = $newReviewCode+"_" + $firstName.Trim()+ "_"+ $lastName.Trim() +"_" + $endDate.Replace("/","-") + "_"+$oldFileName.Replace("_null","")
        Write-Host $newFilename
        Rename-Item -Path $fullFileName -NewName $newFilename

    } elseif ($reviewCode.Contains("FM-verslag")) {
        $newReviewCode =  "FM-verslag"
        $firstName = $text.Substring(($text.IndexOf('First name:') + 11),(($text.IndexOf('Start date:',$text.IndexOf('Start date:')+1) - $text.IndexOf('First name:'))-11))
        $LastName = $text.Substring(($text.IndexOf('Last name:') + 10),(($text.IndexOf('Department:')+1) - $text.IndexOf('Last name:'))-11)
        $endDate = $text.Substring(($text.IndexOf('End Date:') + 10),10)
        $newFilename = $newReviewCode+"_" + $firstName.Trim()+ "_"+ $lastName.Trim() +"_" + $endDate.Replace("/","-") + "_"+$oldFileName.Replace("_null","")
        Write-Host $newFilename
        Rename-Item -Path $fullFileName -NewName $newFilename

    } else {
        $newReviewCode =  "Gespreksnotitie"
        $employeeName = $text.Substring(($text.IndexOf('Employee Name:') + 14),(($text.IndexOf('End Date:') - $text.IndexOf('Employee Name:')-14)))
        $newEmployeeName = $employeeName.Replace(" ","").Trim().Replace(",","_")
        # $LastName = $text.Substring(($text.IndexOf('Last name:') + 10),(($text.IndexOf('Department:')+1) - $text.IndexOf('Last name:'))-11)
        $endDate = $text.Substring(($text.IndexOf('End Date:') + 10),10).Replace("/","-")
        $newFilename = $newReviewCode+"_" + $newEmployeeName+"_"+$oldFileName.Replace("_null","")
        Write-Host $newFilename
        Rename-Item -Path $fullFileName -NewName $newFilename

    }

}
