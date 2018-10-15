$sw = [Diagnostics.Stopwatch]::StartNew()

Function SetupExcel(){ 
    $excelApp = New-Object -ComObject Excel.Application

    $excelApp.visible = $False
    $excelApp.displayalerts = $false
    $excelApp.EnableEvents = $False
    $excelApp.ScreenUpdating = $False
    $excelApp.AskToUpdateLinks = $False
    return $excelApp
}

$folderPath = "FolderOfExcels"
$find = "StringToReplace"
$replace = "ReplaceWith"

$fileNameList = (dir $folderPath).Name
$excelApp =  SetupExcel

Foreach($fileName in $fileNameList){
    
    $workbook = $excelApp.workbooks.open($folderPath +$fileName)
    $workbook.CheckCompatibility = $False
    write-host $workbook.name

    Foreach($worksheet in $workbook.worksheets){
        if($worksheet.name -like "WorksheetName"){
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
