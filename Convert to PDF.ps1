$path = “C:\TimeSheets” 
$xlFixedFormat = “Microsoft.Office.Interop.Excel.xlFixedFormatType” -as [type] 
$excelFiles = Get-ChildItem -File "DM.xlsx" -Path $path 
$objExcel = New-Object -ComObject excel.application 
$objExcel.visible = $false 
foreach($wb in $excelFiles) 
{ 
 


 
$Excel = New-Object -ComObject Excel.Application
$ExcelWorkBook = $Excel.Workbooks.Open($wb.FullName)
$ExcelWorkSheet = $Excel.WorkSheets.item("Demo TS")
$ExcelWorkSheet.activate()
$date = (Get-Date).AddDays(-4).ToString('MM/dd/yyyy')
$ExcelWorkSheet.Cells.Item(6,12) = $date 

$message= "Did you have any vacation days?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
    "Will Have to chose what days."

$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
    "Auto-creates PDF."

$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
switch ($result)
    {
        0 {
      
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Data Entry Form"
$objForm.Size = New-Object System.Drawing.Size(300,260) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {
 $ExcelWorkSheet.Cells.Item(13,1) = $objTextBox.Text
$ExcelWorkSheet.Cells.Item(13,4) = $objTextBox2.Text
$ExcelWorkSheet.Cells.Item(13,7) = $objTextBox3.Text
$ExcelWorkSheet.Cells.Item(13,10) = $objTextBox4.Text
$ExcelWorkSheet.Cells.Item(13,13) = $objTextBox5.Text;
    $objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,165)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({  
#$ExcelWorkSheet.Range("A13:c13").cells(13,1) = "Poop"
$ExcelWorkSheet.Cells.Item(13,1) = $objTextBox.Text
$ExcelWorkSheet.Cells.Item(13,4) = $objTextBox2.Text
$ExcelWorkSheet.Cells.Item(13,7) = $objTextBox3.Text
$ExcelWorkSheet.Cells.Item(13,10) = $objTextBox4.Text
$ExcelWorkSheet.Cells.Item(13,13) = $objTextBox5.Text
    $objForm.Close()})
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,165)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please enter time you worked in space below:"
$objForm.Controls.Add($objLabel) 

$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(10,40) 
$objTextBox.Size = New-Object System.Drawing.Size(50,20) 
$objTextBox.text = "Monday"
$objForm.Controls.Add($objTextBox) 

$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.Location = New-Object System.Drawing.Size(10,65) 
$objTextBox2.Size = New-Object System.Drawing.Size(50,20)
$objTextBox2.text = "Tuesday" 
$objForm.Controls.Add($objTextBox2) 

$objTextBox3 = New-Object System.Windows.Forms.TextBox 
$objTextBox3.Location = New-Object System.Drawing.Size(10,90) 
$objTextBox3.Size = New-Object System.Drawing.Size(50,20) 
$objTextBox3.text = "Wednesday"
$objForm.Controls.Add($objTextBox3) 

$objTextBox4 = New-Object System.Windows.Forms.TextBox 
$objTextBox4.Location = New-Object System.Drawing.Size(10,115) 
$objTextBox4.Size = New-Object System.Drawing.Size(50,20) 
$objTextBox4.text = "Thursday"
$objForm.Controls.Add($objTextBox4) 

$objTextBox5 = New-Object System.Windows.Forms.TextBox 
$objTextBox5.Location = New-Object System.Drawing.Size(10,140) 
$objTextBox5.Size = New-Object System.Drawing.Size(50,20) 
$objTextBox5.text = "Friday"
$objForm.Controls.Add($objTextBox5) 


$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()



                  $ExcelWorkBook.Save()
                 $ExcelWorkBook.Close()
                    $Excel.Quit()

        }
        1 {
        
                  $ExcelWorkBook.Saveas($wb.BaseName,1)
                 $ExcelWorkBook.Close()
                    $Excel.Quit()

        }
    }






$filename = $date -replace "/","."
 
 $filepath = Join-Path -Path $path -ChildPath ($filename + “.pdf”) 
 $workbook = $objExcel.workbooks.open($wb.fullname, 3) 
 $workbook.Saved = $true 
“saving $filepath” 
 $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath) 
 $objExcel.Workbooks.close() 
} 
$objExcel.Quit()