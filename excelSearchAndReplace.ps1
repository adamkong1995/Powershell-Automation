$sw = [Diagnostics.Stopwatch]::StartNew()

Function SetupExcel(){ 
    $excelApp = New-Object -ComObject Excel.Application

    $excelApp.visible = $clear
    $excelApp.displayalerts = $false
    $excelApp.EnableEvents = $False
    $excelApp.ScreenUpdating = $False
    $excelApp.AskToUpdateLinks = $False
    return $excelApp
}

$folderPath = Read-Host -prompt "Input the folder path of excel files, eg: 'O:\Adam\Active Record Testing\'" 
$worksheetName = Read-Host -prompt "Input which worksheet in the workbook to find and replace, eg: 'Data*'" 
$find = Read-Host -prompt "Input the text to find and replace"
$replace = Read-Host -prompt "Input the new text to replace the original text"

$isRun = Read-Host -prompt "
Folder path of the excel files:      $folderPath
Worksheet to find and replace:       $worksheetName
Text to find:                        $find
Text to replace with:                $replace
Please check if the above information is correct (Y/N)"

if($isRun -ne "Y" -and $isRun -ne "N"){
    Write-host "Please input correctly (Y/N) to run the program"
    Exit
}
if($isRun -eq "N"){
    write-host "terminated"
    Exit
}

Write-host "Program start"
$fileNameList = (dir $folderPath).Name
$excelApp =  SetupExcel

Foreach($fileName in $fileNameList){
    
    $workbook = $excelApp.workbooks.open($folderPath +$fileName)
    $workbook.CheckCompatibility = $False
    write-host $workbook.name

    Foreach($worksheet in $workbook.worksheets){
        if($worksheet.name -like $worksheetName){
            write-host $worksheet.name
            $Range = $worksheet.range("A1:AZ500")
            $Search = $Range.find($find)

            if ($search -ne $null) {
                $firstSearchRow = $search.Row
                $firstSearchcolumn = $search.column
        
                $Search.value() = $replace
                $search = $Range.FindNext($search)

                do {
                    $row = $search.row
                    $col = $search.column
                    write-host("Replaced, $col, $Row")
                    $Search.value() = $replace
                    $search = $Range.FindNext($search)
                    $Search.value() = $replace
                } while (($search.row -ne $firstSearchRow) -and ($search.column -ne $firstSearchcolumn))
            }
        }
    }

    $workbook.Save()
    $workbook.Close($True)
    write-host("Close success")
    write-host("")
}

$sw.Stop()
$sw.Elapsed
