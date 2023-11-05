$Excel = New-Object -ComObject Excel.Application

$Excelworkbook1 =
$Excel.Workbooks.Open("C:\Users\hdjom\Downloads\Data_2022.xlsx")
$Excelworkbook2 =
$Excel.Workbooks.Open("C:\Users\hdjom\Downloads\Data_2023.xlsx")

$Sheet1 = $Excelworkbook1.Sheets.Item(1)
$Sheet2 = $Excelworkbook2.Sheets.Item(1)

$UsedRange1 = $Sheet1.UsedRange
$UsedRange2 = $Sheet2.UsedRange

$rowCount = $UsedRange1.Rows.Count
$rowCount2 = $UsedRange2.Rows.Count

$NewWorkbook = $Excel.Workbooks.Add()
$NewSheet = $NewWorkbook.Sheets.Item(1)

$rowA3 = 1

for ($row = 2; $row -le $rowCount; $row++) {
    $ville1 = $Sheet1.Cells.Item($row, 1).Text
    $habitantsA1 = $Sheet1.Cells.Item($row, 2).Value()

    for ($row2 = 2; $row2 -le $rowCount2; $row2++) {
        $ville2 = $Sheet2.Cells.Item($row2, 1).Text
        $habitantsA2 = $Sheet2.Cells.Item($row2, 2).Value()
        
        if($ville1.ToString() -eq $ville2.ToString()){
            
            if ($habitantsA1 -ne $habitantsA2) {
                $NewSheet.Cells.Item($rowA3, 1).value2 = $ville1
                $NewSheet.Cells.Item($rowA3, 2).value2 = ($habitantsA1 - $habitantsA2).ToString()
                $rowA3++
            }
        }
    }

}

$NewWorkbook.SaveAs("C:\Users\hdjom\Downloads\Differences.xlsx")

$Excelworkbook1.Close($false)
$Excelworkbook2.Close($false)
$NewWorkbook.Close($false)
$Excel.Quit()