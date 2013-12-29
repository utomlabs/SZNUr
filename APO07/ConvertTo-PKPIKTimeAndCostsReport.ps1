## .ExternalHelp C:\Users\IK0212141\Documents\APO07\Tools\ConvertTo-PKPIKTimeAndCostsReport.ps1-help.xml
Param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage="Plik z raportem miesięcznym.")]
    [Alias("MonthlyReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$MR,
    
    [Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage="Rok księgowy.")]
    [Alias("Year")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$y,
    
    [Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage="Miesiąc księgowy.")]
    [Alias("Month")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$m
  )
  
Set-StrictMode -Version Latest

Import-LocalizedData -BindingVariable MsgTable

Function Main{
  Clear-History
  Clear-Host
  
  Init-Modules
  $RootLog = Start-LoggerSvc -Configuration "C:\Users\IK0212141\Documents\APO07\Tools\ConvertTo-PKPIKTimeAndCostsReport.ps1.config"  
  $RootLog.Info($MsgTable.StartMsg)
  
  $SrcReportSpec = "Nazwisko i imię","Identyfikator","Pozycja kosztów","Minuty","Godz. min","Kwota"
  $TrgReportSpec = "Identyfikator","Pozycja kosztów","Minuty","Godz. min"
  
  $Delta = $SrcReportSpec
  $TrgReportSpec | %{$Delta = $Delta -ne $_}
  
  Trim-Header -MonthlyReport $MR -AttrSpec $SrcReportSpec | Trim-EmptyRows | Delete-Columns -AttrSpec $Delta | Out-Null
  
  Format-Pivotable -MonthlyReport $MR -AttrSpec @("Identyfikator") | Out-Null #-AttrSpec $TrgReportSpec

  $RootLog.Info($MsgTable.StopMsg)

  Stop-LoggerSvc
  Clear-Modules                                                                                                                                      
}

Function Trim-Header{
Param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Plik z raportem miesięcznym.")]
    [Alias("MonthlyReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$MR,
    
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Specyfikacja nagłówka.")]
    [Alias("AttrSpec")]
    [ValidateNotNullOrEmpty()]
    [Array]$AS
  )
Begin{
  $ScriptLog = Get-Logger -ln Trim-Header
  $ScriptLog.Info($MsgTable.TrimHdrStartMsg)

  Add-Type -AssemblyName Microsoft.Office.Interop.Excel
  $Src = New-Object -ComObject Excel.Application
  }
Process{
  try{
    if(-not $(Test-Path $MR)){
      throw "Monthly report file does not exist!"
    }
    
    $SrcWorkBook = $Src.Workbooks.Open($MR)
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    
    $NumRows = $SrcWorkSheet.UsedRange.Rows.Count
    $NumCols = $AS.Count
    $i = 1
    do{
      $IsHdrRow = $true
      $j = 1
      do{
        $IsHdrRow = $IsHdrRow -and $SrcWorkSheet.UsedRange.Rows.Item($i).Cells.Item($j).Text -eq $AS[$j-1]
        Write-Verbose "i=$i;j=$j;IsHdrRow=$IsHdrRow"
      }while($IsHdrRow -and $j++ -lt $NumCols)
    }while(-not $IsHdrRow -and $i++ -lt $NumRows)
    if(-not $IsHdrRow){
      throw "No header row found!!!"
    }
  }
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Monthly report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        Exit
      }
      "No header row found!!!" {
        $ScriptLog.Error($MsgTable.NoHdrFndMsg)
        Exit
      }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
      }
    }
  }
  catch{
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    Exit
  }
  $HdrRow = $SrcWorkSheet.Rows.Item(1)
  $IsNotFirstRow = $false
  for($i = 0;$i -lt $AS.Count;$i++){$IsNotFirstRow = $IsNotFirstRow -or $HdrRow.Cells.Item($i+1).Text -ne $AS[$i]}
  while($IsNotFirstRow){
    $ScriptLog.Info($MsgTable.RowDelMsg + "$($HdrRow.Cells.Item(1).Text)")
    $HdrRow.Delete() | Out-Null
    $HdrRow = $SrcWorkSheet.Rows.Item(1)
    $IsNotFirstRow = $false
    for($i = 0;$i -lt $AS.Count;$i++){$IsNotFirstRow = $IsNotFirstRow -or $HdrRow.Cells.Item($i+1).Text -ne $AS[$i]}
    }
  $SrcWorkBook.Save()
  $SrcWorkBook.Close()
  $MR
  }
End{
  $Src.Quit() | Out-Null
  ## See http://technet.microsoft.com/en-us/library/ff730962.aspx
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Src) | Out-Null
  Remove-Variable Src
  $ScriptLog.Info($MsgTable.TrimHdrStopMsg)
  }
}

Function Trim-EmptyRows{
Param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Plik z raportem miesięcznym.")]
    [Alias("MonthlyReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$MR
  )
Begin{
  $ScriptLog = Get-Logger -ln Trim-EmptyRows
  $ScriptLog.Info($MsgTable.TrimEmpRowsStartMsg)
  
  Add-Type -AssemblyName Microsoft.Office.Interop.Excel
  $Src = New-Object -ComObject Excel.Application
  }
Process{
  try{
    if(-not $(Test-Path $MR)){
      throw "Monthly report file does not exist!"
    }

    $SrcWorkBook = $Src.Workbooks.Open($MR)
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    }
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Monthly report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        Exit
      }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
      }
    }
  }
  catch{
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    Exit
    }

  $DataRowNum = $SrcWorkSheet.UsedRange.Rows.Count
  $SrcWorkSheet.UsedRange.Rows | %{
    $status = $true
    $_.Columns | %{$status = $status -and [string]::IsNullOrEmpty($_.Cells.Item(1).Text.Trim())}
    if($status){
      $ScriptLog.Info($MsgTable.RowDelMsg + "$($_.Row)")
      $_.Delete() | Out-Null
      }
   Write-Progress -Activity $MsgTable.TrimEmpRowsProcAct -Status $MsgTable.TrimEmpRowsProcStatus -PercentComplete ($_.Row/$DataRowNum*100)
    }
  $SrcWorkBook.Save()
  $SrcWorkBook.Close()
  $MR
  }
End{
  $Src.Quit() | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Src) | Out-Null
  Remove-Variable Src
  $ScriptLog.Info($MsgTable.TrimEmpRowsStopMsg)
  }
}

Function Delete-Columns{
Param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Plik z raportem miesięcznym.")]
    [Alias("MonthlyReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$MR,
    
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Specyfikacja kolumn do usunięcia.")]
    [Alias("AttrSpec")]
    [ValidateNotNullOrEmpty()]
    [Array]$AS
  )
Begin{
  $ScriptLog = Get-Logger -ln Delete-Columns
  $ScriptLog.Info($MsgTable.DelColsStartMsg)
  
  Add-Type -AssemblyName Microsoft.Office.Interop.Excel
  $Src = New-Object -ComObject Excel.Application
  }
Process{
  try{
    if(-not $(Test-Path $MR)){
      throw "Monthly report file does not exist!"
    }

    $SrcWorkBook = $Src.Workbooks.Open($MR)
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    }
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Monthly report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        Exit
      }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
      }
    }
  }
  catch {
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    Exit
    }

  $AS | %{
    $ScriptLog.Info($MsgTable.ColDelMsg + "$($SrcWorkSheet.UsedRange.Rows.Item(1).Find($_).EntireColumn.Column)")
    $SrcWorkSheet.UsedRange.Rows.Item(1).Find($_).EntireColumn.Delete() | Out-Null
  }

  $SrcWorkBook.Save()
  $SrcWorkBook.Close()
  $MR
  }
End{
  $Src.Quit() | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Src) | Out-Null
  Remove-Variable Src
  $ScriptLog.Info($MsgTable.DelColsStopMsg)
  }
}

Function Format-Pivotable{
Param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Plik z raportem miesięcznym.")]
    [Alias("MonthlyReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$MR,
    
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Specyfikacja kolumn raportu.")]
    [Alias("AttrSpec")]
    [ValidateNotNullOrEmpty()]
    [Array]$AS
  )
Begin{
  $ScriptLog = Get-Logger -ln Format-Pivotable
  $ScriptLog.Info($MsgTable.FormPivotStartMsg)

  Add-Type -AssemblyName Microsoft.Office.Interop.Excel
  $Src = New-Object -ComObject Excel.Application
  }
Process{
  try{
    if(-not $(Test-Path $MR)){
      throw "Monthly report file does not exist!"
    }

    $SrcWorkBook = $Src.Workbooks.Open($MR)
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    }
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Monthly report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        Exit
      }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
      }
    }
  }
  catch {
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    Exit
    }
  
  $xlSrcType = [Microsoft.Office.Interop.Excel.XlListObjectSourceType]
  $xlYesNo = [Microsoft.Office.Interop.Excel.XlYesNoGuess]
  
  $SrcListObject = $SrcWorkSheet.ListObjects.Add($xlSrcType::xlSrcRange,$SrcWorkSheet.UsedRange, $null,$xlYesNo::xlYes) #xlSrcRange, xlYes
  $SrcListObject.DisplayName = "SourceData"
  $SrcListObject.Name = "SourceData"

  $ValRows = New-Object System.Collections.ArrayList
  $BufferLen = 2
  $Buffer = New-Object System.Collections.ArrayList -ArgumentList $BufferLen
  $AS | %{
    #Following section is a candidate for routine Find-Values. The most inner if statement has to be generalised
    $DataRowNum = $SrcWorkSheet.Range("SourceData[$_]").Rows.Count
    $SrcWorkSheet.Range("SourceData[$_]") | %{
      if($Buffer.Count -lt $BufferLen) { $Buffer.Add([boolean]$_.Text.Trim()) | Out-Null } else { $Buffer.RemoveAt(0); $Buffer.Add([boolean]$_.Text.Trim()) | Out-Null }
      #Following line should be generalised to not only first two buffer values, but patterns
      if($Buffer.Count -eq $BufferLen -and $Buffer[0] -and !$Buffer[1]) {
        $ValRows.Add($($_.Row - $BufferLen + 1)) | Out-Null
        }
      Write-Progress -Activity $MsgTable.FormPivotFindValProcAct -Status $MsgTable.FormPivotFindValProcStatus -PercentComplete ($_.Row/$DataRowNum*100)
      }
    #Following section fills missing values
    for($i = 0;$i -lt $ValRows.Count-1;$i++){
      $SrcWorkSheet.Range("SourceData[$_]").Range("A$($ValRows[$i]-1):A$($ValRows[$i+1]-2)").FillDown() | Out-Null
      }
    $SrcWorkSheet.Range("SourceData[$_]").Range("A$($ValRows[$ValRows.Count-1]-1):A$($DataRowNum)").FillDown() | Out-Null
    }
  
  $SrcWorkBook.Save()
  $SrcWorkBook.Close()
  $MR
  }
End{
  $Src.Quit() | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Src) | Out-Null
  Remove-Variable Src
  $ScriptLog.Info($MsgTable.FormPivotStopMsg)
  }
}

Function Init-Modules{
  try{
    $Script:ModList = New-Object System.Collections.ArrayList
    $ModList.Add(@(Import-Module -Name PSLog -ArgumentList "C:\Users\IK0212141\Documents\WindowsPowerShell\Libs\log4net\bin\net\3.5\release\log4net.dll" -Force -PassThru)) | Out-Null
  }
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Log4net library cannot be found on the path" {
        Write-Error $MsgTable.Log4NetPathMsg
      }
      default {
        Write-Error $MsgTable.DefaultNegMsg
      }
    }
  }
  catch{
    "*"*80
    $_.Exception.GetType().FullName
    $_.Exception.Message
    "*"*80
    Exit
  }
}

Function Clear-Modules{
  $ModList | %{Remove-Module $_}
}

. Main

# SIG # Begin signature block
# MIIEMwYJKoZIhvcNAQcCoIIEJDCCBCACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUaZUbd7ssu/ivGvjT9TG0+MS9
# zqagggI9MIICOTCCAaagAwIBAgIQtBirZz3Acb1BfUstCv49PTAJBgUrDgMCHQUA
# MCwxKjAoBgNVBAMTIVBvd2VyU2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdDAe
# Fw0xMzA1MDYyMjExMTJaFw0zOTEyMzEyMzU5NTlaMBoxGDAWBgNVBAMTD1Bvd2Vy
# U2hlbGwgVXNlcjCBnzANBgkqhkiG9w0BAQEFAAOBjQAwgYkCgYEAitSmlTAETOy4
# uI7gmQoTK8zKCb8VEStM9gqQtVxcO2HfEgpTnl8NbcXsqwfjiRvQ+qUpJyO6dBaM
# /DU8ZxtKn4bBRofjMiYTH1VLqIDZweqHLQQFAmV9tKB28L9JxZKROqnuW6rD3+u1
# BGKdOEViA9ogRmDTif7evlloDHeKFqsCAwEAAaN2MHQwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwXQYDVR0BBFYwVIAQAT6NGGMwu5QiCSwIlq1wnaEuMCwxKjAoBgNVBAMT
# IVBvd2VyU2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdIIQ1iyEzXrW9apItH1h
# a/owUTAJBgUrDgMCHQUAA4GBADv9uxMjxKwJzPtNjakjYKLVEFxujzkbs51SK/yb
# 1LamnYdJ7pgFYhsZH+6aRlC06V0CGlAnBvXlUksj289x/BLE3osm7Xc9UJBqrUXu
# B8svNR4vHgjs5GBqcFNtVe0xm5YVlCTzfTBNhpdO+W3HpEUZhf046Wgl+bJErIRH
# SEKDMYIBYDCCAVwCAQEwQDAsMSowKAYDVQQDEyFQb3dlclNoZWxsIExvY2FsIENl
# cnRpZmljYXRlIFJvb3QCELQYq2c9wHG9QX1LLQr+PT0wCQYFKw4DAhoFAKB4MBgG
# CisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
# AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYE
# FJMO9DhcCeeYx68G7zAm5HohD/0bMA0GCSqGSIb3DQEBAQUABIGAVUtV+XNrQVKv
# 4AoKj7FxxIT0JnQczoMS5NL9LDpb6QCeweAluWJDkRBzzxMCkDHTMgYMwhbl8Bb/
# jLP5LNbP0paHrNxSDnI8dIqGOXnIYBQR87AtsrGTBd3Q/rdwZH1RYg4vvWY0Dbyv
# bxwWH8BbfyKacL5yfJjghyt3SfpuZUs=
# SIG # End signature block
