## .ExternalHelp C:\Users\IK0212141\Documents\APO07\Tools\ConvertTo-PKPIKTimeAndCostsReport.ps1-help.xml
Param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage="Plik z raportem miesięcznym.")]
    [Alias("MonthlyReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$MR,
    
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage="Plik z raportem skonsolidowanym.")]
    [Alias("ConsolidatedReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$CR,
    
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
  
  #Trim-Description -MonthlyReport $MR -AttrSpec $SrcReportSpec | Trim-EmptyRows | Delete-Columns -AttrSpec $Delta | Out-Null
  
  #Format-Pivotable -MonthlyReport $MR -AttrSpec @("Identyfikator") | Out-Null #-AttrSpec $TrgReportSpec
  
  Import-Report -SourceReport $MR -TargetReport $CR -AttrSpec $TrgReportSpec -Year $y -Month $m

  $RootLog.Info($MsgTable.StopMsg)

  Stop-LoggerSvc
  Clear-Modules                                                                                                                                      
}

Function Trim-Description{
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
  $ScriptLog = Get-Logger -ln Trim-Description
  $ScriptLog.Info($MsgTable.TrimDescStartMsg)

  Add-Type -AssemblyName Microsoft.Office.Interop.Excel
  }
Process{
  try{
    $SrcWorkBook = Open-Workbook -Report $MR
    
    Find-Header -MonthlyReportWorkbook $SrcWorkBook -AttrSpec $AS
    
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    $DescRow = $SrcWorkSheet.Rows.Item(1)
    $IsNotFirstRow = $false
    for($i = 0;$i -lt $AS.Count;$i++){$IsNotFirstRow = $IsNotFirstRow -or $DescRow.Cells.Item($i+1).Text -ne $AS[$i]}
    while($IsNotFirstRow){
      $ScriptLog.Info($MsgTable.RowDelMsg + "$($DescRow.Cells.Item(1).Text)")
      $DescRow.Delete() | Out-Null
      $DescRow = $SrcWorkSheet.Rows.Item(1)
      $IsNotFirstRow = $false
      for($i = 0;$i -lt $AS.Count;$i++){$IsNotFirstRow = $IsNotFirstRow -or $DescRow.Cells.Item($i+1).Text -ne $AS[$i]}
      }
    $SrcWorkBook.Save()
    $SrcWorkBook.Close()
    $MR
    }
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        }
      "No header row found!!!" {
        $ScriptLog.Error($MsgTable.NoHdrFoundMsg)
        }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
        $ScriptLog.Error($_)
        }
      }
    }
  catch{
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    }
  finally{
    if($SrcWorkBook.Application){
      $SrcWorkBook.Application.Quit() | Out-Null
      }
    }
  }
End{
  ## See http://technet.microsoft.com/en-us/library/ff730962.aspx
  ##[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Src) | Out-Null
  ##Remove-Variable Src
  $ScriptLog.Info($MsgTable.TrimDescStopMsg)
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
  }
Process{
  try{
    $SrcWorkBook = Open-Workbook -Report $MR
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    
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
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
        $ScriptLog.Error($_)
        }
      }
    }
  catch{
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    }
  finally{
    if($SrcWorkBook.Application){
      $SrcWorkBook.Application.Quit() | Out-Null
      }
    }
  }
End{
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
  }
Process{
  try{
    $SrcWorkBook = Open-Workbook -Report $MR
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    $AS | %{
      $ScriptLog.Info($MsgTable.ColDelMsg + "$($SrcWorkSheet.UsedRange.Rows.Item(1).Find($_).EntireColumn.Column)")
      $SrcWorkSheet.UsedRange.Rows.Item(1).Find($_).EntireColumn.Delete() | Out-Null
      }
    $SrcWorkBook.Save()
    $SrcWorkBook.Close()
    $MR
    }
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
        $ScriptLog.Error($_)
        }
      }
    }
  catch {
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    }
  finally{
    if($SrcWorkBook){
      $SrcWorkBook.Application.Quit() | Out-Null
      }
    }
  }
End{
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
  }
Process{
  try{
    $SrcWorkBook = Open-Workbook -Report $MR
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    
    $xlSrcType = [Microsoft.Office.Interop.Excel.XlListObjectSourceType]
    $xlYesNo = [Microsoft.Office.Interop.Excel.XlYesNoGuess]
    
    if($SrcWorkSheet.ListObjects | ?{$_.Name -eq "SourceData"}){throw "Pivot ready list already exists!"}
    
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
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        }
      "Pivot ready list already exists!" {
        $ScriptLog.Error($MsgTable.PivotblListExistMsg)
      }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
        $ScriptLog.Error($_)
        }
      }
    }
  catch {
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    }
  finally{
    if($SrcWorkBook.Application){
      $SrcWorkBook.Application.Quit() | Out-Null
      }
    }
  }
End{
  $ScriptLog.Info($MsgTable.FormPivotStopMsg)
  }
}

Function Import-Report{
param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Plik z raportem miesięcznym.")]
    [Alias("SourceReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$SR,
    
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage="Plik z raportem skonsolidowanym.")]
    [Alias("TargetReport")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$TR,
    
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage="Specyfikacja nagłówka.")]
    [Alias("AttrSpec")]
    [ValidateNotNullOrEmpty()]
    [Array]$AS,

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
Begin{
  $ScriptLog = Get-Logger -ln Import-Report
  $ScriptLog.Info($MsgTable.ImportRepStartMsg)

  Add-Type -AssemblyName Microsoft.Office.Interop.Excel
  }
Process{
  try{
    $SrcWorkBook = Open-Workbook -Report $SR
    $SrcWorkSheet = $SrcWorkBook.Worksheets.Item(1)
    
    $TrgWorkBook = Open-Workbook -Report $TR
    $TrgWorkSheet = $TrgWorkBook.Worksheets.Item("Źródło")
    
    $SrcRowNum = $SrcWorkSheet.Range("SourceData[#Dane]").Rows.Count
    $SrcWorkSheet.Range("SourceData[#Dane]").Rows | %{
      $TempListRow = $TrgWorkSheet.ListObjects.Item("Dane").ListRows.Add()
      $TempListRow.Range.Cells.Item(1).Value2 = $y
      $TempListRow.Range.Cells.Item(2).Value2 = $m
      $TempListRow.Range.Cells.Item(3).Value2 = $_.Cells.Item(1).Value2
      $TempListRow.Range.Cells.Item(4).Value2 = $_.Cells.Item(2).Value2
      $TempListRow.Range.Cells.Item(5).Formula = "=WYSZUKAJ.PIONOWO(FRAGMENT.TEKSTU([@[Pozycja kosztów]];17;3);'C:\Users\IK0212141\Documents\APO06\ZPK.xlsx'!MPK_MPP[#Dane];2;FAŁSZ)"
      $TempListRow.Range.Cells.Item(6).Formula = "=WYSZUKAJ.PIONOWO([@Identyfikator];'C:\Users\IK0212141\Documents\APO07\Arkusz_osobowy_lok_sorg_spr.xlsx'!LokSorgSpr$y$m[#Dane];7;FAŁSZ)"
      $TempListRow.Range.Cells.Item(7).Formula = "=WYSZUKAJ.PIONOWO([@Identyfikator];'C:\Users\IK0212141\Documents\APO07\Arkusz_osobowy_lok_sorg_spr.xlsx'!LokSorgSpr$y$m[#Dane];8;FAŁSZ)"
      $TempListRow.Range.Cells.Item(8).Formula = "=WYSZUKAJ.PIONOWO([@Identyfikator];'C:\Users\IK0212141\Documents\APO07\Arkusz_osobowy_lok_sorg_spr.xlsx'!LokSorgSpr$y$m[#Dane];9;FAŁSZ)"
      $TempListRow.Range.Cells.Item(9).Formula = "=WYSZUKAJ.PIONOWO(FRAGMENT.TEKSTU([@[Pozycja kosztów]];21;3);'C:\Users\IK0212141\Documents\APO06\ZPK.xlsx'!PROCES[#Dane];2;FAŁSZ)"
      $TempListRow.Range.Cells.Item(10).Formula = "=WYSZUKAJ.PIONOWO(FRAGMENT.TEKSTU([@[Pozycja kosztów]];25;3);'C:\Users\IK0212141\Documents\APO06\ZPK.xlsx'!PRODUKT[#Dane];2;FAŁSZ)"
      $TempListRow.Range.Cells.Item(11).Formula = "=WYSZUKAJ.PIONOWO(FRAGMENT.TEKSTU([@[Pozycja kosztów]];29;3);'C:\Users\IK0212141\Documents\APO06\ZPK.xlsx'!KLIENT_WEW[#Dane];2;FAŁSZ)"
      $TempListRow.Range.Cells.Item(12).Formula = "=WYSZUKAJ.PIONOWO(FRAGMENT.TEKSTU([@[Pozycja kosztów]];33;2);'C:\Users\IK0212141\Documents\APO06\ZPK.xlsx'!DZIAŁALNOŚĆ[#Dane];2;FAŁSZ)"
      $TempListRow.Range.Cells.Item(13).Formula = "=WYSZUKAJ.PIONOWO(FRAGMENT.TEKSTU([@[Pozycja kosztów]];36;2);'C:\Users\IK0212141\Documents\APO06\ZPK.xlsx'!KLIENT_ZEW[#Dane];2;FAŁSZ)"
      $TempListRow.Range.Cells.Item(14).Value2 = $_.Cells.Item(3).Value2
      $TempListRow.Range.Cells.Item(15).Value2 = $_.Cells.Item(4).Value2

      Write-Progress -Activity $MsgTable.ImportRepProcAct -Status $MsgTable.ImportRepProcStatus -PercentComplete ($_.Row/$SrcRowNum*100)
      }
    
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()

    $TrgWorkBook.Save()
    $TrgWorkBook.Close()
    
    $TR
    }
  catch [System.Management.Automation.RuntimeException] {
    switch($_.Exception.Message){
      "Report file does not exist!" {
        $ScriptLog.Error($MsgTable.NoMonthRepMsg)
        }
      default {
        $ScriptLog.Error($MsgTable.DefaultNegMsg)
        $ScriptLog.Error($_)
        }
      }
    }
  catch {
    $ScriptLog.Error("$($_.Exception.GetType().FullName)`n$($_.Exception.Message)")
    
    $SrcWorkBook.Saved = $true
    $SrcWorkBook.Close()
    
    $TrgWorkBook.Saved = $true
    $TrgWorkBook.Close()
    }
  finally{
    if($SrcWorkBook.Application){
      $SrcWorkBook.Application.Quit() | Out-Null
      }
    if($TrgWorkBook.Application){
      $TrgWorkBook.Application.Quit() | Out-Null
      }
    }  
  }
End{
  $ScriptLog.Info($MsgTable.ImportRepStopMsg)
  }
}

Function Open-Workbook{
param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Plik z raportem miesięcznym.")]
    [Alias("Report")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$R
  )
  Add-Type -AssemblyName Microsoft.Office.Interop.Excel
  $Src = New-Object -ComObject Excel.Application
  
  if(-not $(Test-Path $R)){
    throw "Report file does not exist!"
  }
  
  $Src.Workbooks.Open($R)
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

Function Find-Header{
param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage="Skoroszyt z raportem miesięcznym.")]
    [Alias("MonthlyReportWorkbook")]
    [ValidateNotNullOrEmpty()]
    [System.__ComObject]$MRW,
    
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Specyfikacja nagłówka.")]
    [Alias("AttrSpec")]
    [ValidateNotNullOrEmpty()]
    [Array]$AS
  )  
  $SrcWorkSheet = $MRW.Worksheets.Item(1)
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
    Write-Progress -Activity $MsgTable.FindHdrProcAct -Status $MsgTable.FindHdrProcStatus -PercentComplete ($i/$NumRows*100)
  }while(-not $IsHdrRow -and $i++ -lt $NumRows)
  
  if(-not $IsHdrRow){
    throw "No header row found!!!"
  }
}

Function Clear-Modules{
  $ModList | %{Remove-Module $_}
}

. Main

# SIG # Begin signature block
# MIIEMwYJKoZIhvcNAQcCoIIEJDCCBCACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU1hF8QE+RP2IG2mRHJFlGiiJI
# WaegggI9MIICOTCCAaagAwIBAgIQtBirZz3Acb1BfUstCv49PTAJBgUrDgMCHQUA
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
# FMsTj9NrAm7yfBowJD+otTqg+GETMA0GCSqGSIb3DQEBAQUABIGAFL9ayqFN3sTW
# KXZfF9nG1C+r3lkD+/9kSMn+cUNKNdpMc5IRAl6jDT4KtBuWFJum8vfSppQiBizG
# 7n4mjwsZ6v8/YqaKiYKzaXi2DqtxpqS1OluLr4uLnuxPSs2akp/y6Y9ZSTn7j2bV
# 1Zrh0byNYN70Qde2baHi7ImpMpOFSV4=
# SIG # End signature block
