##############################
## Powershell - Excel Split ##
##############################
# Script for split in single file an Excel List by filter
#Define local variables
$strRootPath = Split-Path -Path $MyInvocation.MyCommand.Path

function KillExcelProcesses {   
    try
    { 
        # Get processes list
        $arrProcessesList = Get-Process -Name "EXCEL" -ErrorAction 'Stop' | Where-Object {$_.mainWindowTitle -eq ""}
    
        # Cycle for each process EXCEL without MainTitle (no visible)
        foreach($process in $arrProcessesList)
        {
            # Kill process
            try
            {
                $process.kill()    
            }
            catch
            {
            
            }   
        }
    }
    catch [System.Management.Automation.ActionPreferenceStopException]
    {        
    }

    # Sleep for avoid problems
    Start-Sleep -Seconds 3
}

function SplitExcelFile {
    param(
        [Parameter(Mandatory)]
        [string]$FileToSplit,

        [Parameter(Mandatory)]
        [string]$ColumnToFilter
    )

    # Define variables
    $strFileExcel = ("{0}\FileSorgenti\{1}" -f $strRootPath, $FileToSplit)
    $strOutPath = ("{0}\DestinazioneFile" -f $strRootPath)
    $arrColExcel = New-Object System.Collections.ArrayList
    $typMissing = [Type]::Missing

    # Open source file Excel and set visible and caption proprieties and disable screen messages
    # N.B. set visible for avoid to kill process on array cycle for values
    $objSouExcel = New-Object -ComObject "Excel.Application"
    $objSouExcel.Visible = $true
    $objSouExcel.Caption = "MasterExcelProcess"
    $objSouExcel.DisplayAlerts = $false

    # Open Workbook
    $objSouWorkbook = $objSouExcel.Workbooks.Open($strFileExcel)

    # Get object sheet as sheet 1 (first sheet file excel) and activate it
    $objSouWorkSheet = $objSouWorkbook.Sheets.Item(1) 
    $objSouWorkSheet.Activate() | Out-Null

    # Get rows count and columns count
    $intLastRow = $objSouWorkSheet.UsedRange.Rows.count
    $intLastCol = $objSouWorkSheet.UsedRange.Columns.count
    
    # Get column index of column used to filter
    $intColIdx = $objSouWorkSheet.Range(("{0}1" -f $ColumnToFilter)).Column

    # Get object Range, is range for select column to filter and set advanced filter
    $objRngColFilter = $objSouWorkSheet.Range(("{0}1:{0}{1}" -f $ColumnToFilter,$intLastRow))
    $objRngColFilter.AdvancedFilter([Microsoft.Office.Interop.Excel.XlFilterAction]::xlFilterInPlace, $typMissing, $typMissing, $true) | Out-Null

    # Get range for visible cells after filter and get only column used for split file
    $objRngForGetValues = $objSouWorkSheet.Range(("{0}2:{0}{1}" -f $ColumnToFilter, $intLastRow)).SpecialCells([Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeVisible)
    
    # Cycle rows and add values to array
    foreach($row in $objRngForGetValues.Rows)
    {
        $arrColExcel.Add($row.Value2) | Out-Null
    }

    # Reset filter
    #$objSouWorkSheet.ShowAllData  | Out-Null
    
    # Cycle for every row in array (are the values for split Excel list)
    foreach($value in $arrColExcel)
    {
        # Reset filter
        $objSouWorkSheet.ShowAllData | Out-Null

        # Create Excel temp object for destination file and set proprieties
        $objDstExcelTmp = New-Object -ComObject "Excel.Application"
        $objDstExcelTmp.Visible = $false
        $objDstExcelTmp.DisplayAlerts = $false
    
        # Add Workbook to new file
        $objDstWkTmp = $objDstExcelTmp.Workbooks.Add()

        # Get first sheet and activate it
        $objDstShtmp = $objDstWkTmp.Sheets.Item(1)
        $objDstShtmp.Activate() | Out-Null
        
        # Get source range and filter it with column value
        $objSouWorkSheet.Range($objSouWorkSheet.Cells(1,1), $objSouWorkSheet.Cells($intLastRow,$intLastCol)).AutoFilter($intColIdx,$value) | Out-Null
        
        # Copy used range    
        $objSouWorkSheet.UsedRange.SpecialCells(12).Copy() | Out-Null
        
        # Sleep for avoid problems 
        Start-Sleep -Seconds 2

        # Select range to paste values
        $rngDest = $objDstShtmp.Range("A1") 
        
        # Paste values   
        $objDstShtmp.Paste($rngDest) | Out-Null
       
        # Save file (replace special characters)
        $objDstWkTmp.SaveAs(("{0}\{1}.xlsx" -f $strOutPath, $value.tostring().Replace(".","_").Replace("/","_"))) 

        # Close file and quit Excel
        $objDstWkTmp.Close($false)
        $objDstExcelTmp.Quit() 
        
        # Print success message
        Write-Host ("File per valore {0} creato" -f $value) 

        # Sleep for avoid problems
        Start-Sleep -Seconds 3
    
        # Call function for kill processes
        KillExcelProcesses         
    }

    # Reset filter
    $objSouWorkSheet.ShowAllData | Out-Null

    # Close workbook and quit excel
    $objSouWorkbook.Close($false)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objSouWorkbook) | Out-Null
    $objSouExcel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objSouExcel) | Out-Null

    # Call function for kill processes
    KillExcelProcesses 
}


# Read files into source folder (only Excel file)
$arrFiles = Get-ChildItem -Path ("{0}\FileSorgenti" -f $strRootPath) | Where-Object {$_.Extension -eq ".xlsx" -or $_.Extension -eq ".xls"} | Select Name

# Cycle for each file
foreach($objItem in $arrFiles)
{    
    # Interactive parameters
    $inColumnToFilter = Read-Host ("Colonna per cui filtrare il file {0} (lettera, es: A, B)" -f $objItem.Name)
    
    # Call method for split Excel file
    SplitExcelFile -FileToSplit $objItem.Name -ColumnToFilter $inColumnToFilter    
}
