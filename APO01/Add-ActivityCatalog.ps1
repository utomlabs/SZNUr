## .ExternalHelp C:\Users\IK0212141\Documents\WindowsPowerShell\SZNUr\APO01\Add-ActivityCatalog.ps1-help.xml

[CmdletBinding(SupportsShouldProcess=$true)]
Param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage="W jakim pliku znajduje się źródłowy regulamin organizacyjny?")]
    [Alias("rof")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$ROFile,
    
    [Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage="Do jakiego pliku zapisać ekstrakt?")]
    [Alias("of")]
    [ValidateLength(1,254)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
    [String]$OutputFile
  )
  
Set-StrictMode -Version Latest

Import-LocalizedData -BindingVariable MsgTable

function Main
{   
    #[void] $ModList.Add(@(Import-Module -Name PSClass -Force -PassThru))
    . C:\Users\IK0212141\Documents\WindowsPowerShell\Modules\PSClass\PSClass.ps1

    $temp = Get-ActivityCatalog -ROFile $ROFile
    Write-Verbose "Wewnątrz procedury Main wywołanie GetType(): $($temp.GetType())"
    Write-Verbose "Wewnątrz procedury Main wywołanie GetType() dla pierwszego elementu: $($temp[0].GetType())"
    Set-StrictMode -Off
    Write-Verbose "Wewnątrz procedury Main wyświetlenie pierwszego elementu: $($temp[0].ToString())"
    Set-StrictMode -Version Latest    
}

function Get-ActivityCatalog
{
  [CmdletBinding(SupportsShouldProcess=$true)]
  Param
    (
      [Parameter(Mandatory=$true, ValueFromPipeline=$false, HelpMessage="W jakim pliku znajduje się źródłowy regulamin organizacyjny?")]
      [Alias("rof")]
      [ValidateLength(1,254)]
      [ValidateNotNullOrEmpty()]
      [ValidatePattern("^\S([a-z]|[A-Z]|[0-9]|\.|-|_)*")]
      [String]$ROFile
    )
  Begin{
      Clear-History
      Clear-Host

      try{
          $ModList = New-Object System.Collections.ArrayList
          [void] $ModList.Add(@(Import-Module -Name PSLog -ArgumentList "C:\Users\IK0212141\Documents\WindowsPowerShell\Libs\log4net\bin\net\3.5\release\log4net.dll" -Force -PassThru))
          $RootLog = Start-LoggerSvc -Configuration "C:\Users\IK0212141\Documents\WindowsPowerShell\SZNUr\APO01\Add-ActivityCatalog.ps1.config"
          
          #[void] $ModList.Add(@(Import-Module -Name PSClass -Force -PassThru))
          . C:\Users\IK0212141\Documents\WindowsPowerShell\Modules\PSClass\PSClass.ps1
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
      catch {
          "*"*80
          $_.Exception.GetType().FullName
          $_.Exception.Message
          "*"*80
          Exit
      }

      $ActivityClass = New-PSClass Activity {
          note -static ObjectCount 0
          note -private _Domain
          note -private _Process
          note -private _Activity
          note -private _Details
          
          constructor {
              param ($dom,$proc,$act,$det)
              $private._Domain = $dom
              $private._Process = $proc
              $private._Activity = $act
              $private._Details = $det
      
              $ActivityClass.ObjectCount += 1
          }
          
          property Domain {
              $private._Domain
          } -set {
              param($newDomain)
              Write-Verbose "Renaming $($this.Class.ClassName) '$($private._Domain)' to '$($newDomain)'"
              $private._Domain = $newDomain
          }
          
          property Process {
              $private._Process
          } -set {
              param($newProcess)
              Write-Verbose "Renaming $($this.Class.ClassName) '$($private._Process)' to '$($newProcess)'"
              $private._Process = $newProcess
          }
          
          property Activity {
              $private._Activity
          } -set {
              param($newActivity)
              Write-Verbose "Renaming $($this.Class.ClassName) '$($private._Activity)' to '$($newActivity)'"
              $private._Activity = $newActivity
          }
          
          property Details {
              $private._Details
          } -set {
              param($newDetails)
              Write-Verbose "Renaming $($this.Class.ClassName) '$($private._Details)' to '$($newDetails)'"
              $private._Details = $newDetails
          }
          
          method -override ToString {
              "$($this.Class.ClassName);$($this.Domain);$($this.Process);$($this.Activity);$($this.Details)"
          }
          
          method -static DisplayObjectCount {
              "$($this.ClassName) has $($this.ObjectCount) instances"
          }
      }
      
      $RootLog.Info($MsgTable.StartMsg)
      
      $ScriptLog = Get-Logger -ln Add-ActivityCatalog

      ## As Microsoft.ACE.OLEDB.12.0 provider works only on x86 architecture the following statement acts as hardening of the basic functional code
      $ScriptLog.Info($env:Processor_Architecture)
      if ($env:Processor_Architecture -ne "x86") {
          $ScriptLog.Warn("Running x86 PowerShell...")
          $RootLog.Info($MsgTable.StopMsg)
          Stop-LoggerSvc
          if ($script:myInvocation.Line){
              #&"$env:WINDIR\syswow64\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $Script:$myInvocation.Line
              &"$env:WINDIR\syswow64\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile -Command $Script:PSCmdlet.MyInvocation.Line
          } else {
              &"$env:WINDIR\syswow64\windowspowershell\v1.0\powershell.exe" -NonInteractive -File "$($myInvocation.InvocationName)" $args
          }
          exit $lastExitCode
      }
      
      $ActList = New-Object System.Collections.ArrayList
      
      $objOleDbConnection = New-Object "System.Data.OleDb.OleDbConnection"
      $objOleDbCommand = New-Object "System.Data.OleDb.OleDbCommand"
      $objOleDbAdapter = New-Object "System.Data.OleDb.OleDbDataAdapter"
      $objDomainsDataTable = New-Object "System.Data.DataTable"
      $objProcessesDataTable = New-Object "System.Data.DataTable"
      $objActivitiesDataTable = New-Object "System.Data.DataTable"
      $objDetailsDataTable = New-Object "System.Data.DataTable"

      ##Note that only .xls file are supported with JET, 
      ##.xlsx require the Microsoft.ACE provider which is not installed by default.
      ##Also, this only works when run as a 32-bit process on 64-bit operating systems.
      ##Examples: http://www.connectionstrings.com/excel

      ##$objOleDbConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=$ROFile;Extended Properties=""Excel 8.0;HDR=YES"""
      $objOleDbConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=$ROFile; Extended Properties=""Excel 12.0;HDR=YES"""
      $objOleDbConnection.Open()

      #Write-Verbose $objOleDbConnection.State

      $objOleDbCommand.Connection = $objOleDbConnection
      $objOleDbCommand.CommandText = "SELECT DISTINCT Domena FROM [RO-Src$]"
      $objOleDbAdapter.SelectCommand = $objOleDbCommand
      [void] $objOleDbAdapter.Fill($objDomainsDataTable)
      
      ForEach($Domain in $objDomainsDataTable){
          $ScriptLog.Debug("$($Domain.ItemArray)")
          
          ##"SELECT DISTINCT Proces FROM [RO-Src$] WHERE Domena="""+$Domain.ItemArray+""""
          $objOleDbCommand.CommandText = "SELECT DISTINCT Proces FROM [RO-Src$] WHERE Domena=?"
          $objOleDbCommand.Parameters.Add("@Domena",[System.Data.OleDb.OleDbType]::VarChar,256).Value = $($Domain.ItemArray)
          $objOleDbAdapter.SelectCommand = $objOleDbCommand
          [void] $objOleDbAdapter.Fill($objProcessesDataTable)
          $objOleDbCommand.Parameters.Clear()
          
          ForEach($Process in $objProcessesDataTable){
              $ScriptLog.Debug("|->$($Process.ItemArray)")
              
              ##"SELECT DISTINCT Działanie FROM [RO-Src$] WHERE Domena="""+$Domain.ItemArray+""" AND Proces="""+$Process.ItemArray+""""
              $objOleDbCommand.CommandText = "SELECT DISTINCT Działanie FROM [RO-Src$] WHERE Domena=? AND Proces=?"
              $objOleDbCommand.Parameters.Add("@Domena",[System.Data.OleDb.OleDbType]::VarChar,256).Value = $($Domain.ItemArray)
              $objOleDbCommand.Parameters.Add("@Proces",[System.Data.OleDb.OleDbType]::VarChar,256).Value = $($Process.ItemArray)
              $objOleDbAdapter.SelectCommand = $objOleDbCommand
              [void] $objOleDbAdapter.Fill($objActivitiesDataTable)
              $objOleDbCommand.Parameters.Clear()
              
              ForEach($Activity in $objActivitiesDataTable){
                  $ScriptLog.Debug("|--->$($Activity.ItemArray)")
                  $objOleDbCommand.CommandText = "SELECT DISTINCT Uszczegółowienie FROM [RO-Src$] WHERE Domena=? AND Proces=? AND Działanie=?"
                  $objOleDbCommand.Parameters.Add("@Domena",[System.Data.OleDb.OleDbType]::VarChar,256).Value = $($Domain.ItemArray)
                  $objOleDbCommand.Parameters.Add("@Proces",[System.Data.OleDb.OleDbType]::VarChar,256).Value = $($Process.ItemArray)
                  $objOleDbCommand.Parameters.Add("@Działanie",[System.Data.OleDb.OleDbType]::VarChar,256).Value = $($Activity.ItemArray)
                  $objOleDbAdapter.SelectCommand = $objOleDbCommand
                  [void] $objOleDbAdapter.Fill($objDetailsDataTable)
                  $objOleDbCommand.Parameters.Clear()
                  
                  ForEach($Detail in $objDetailsDataTable){
                      $ScriptLog.Debug("|----->$($Detail.ItemArray)")
                      Set-StrictMode -Off
                      [void] $ActList.Add($ActivityClass.New("$($Domain.ItemArray)","$($Process.ItemArray)","$($Activity.ItemArray)","$($Detail.ItemArray)"))
                      Set-StrictMode -Version Latest
                  }
                  $objDetailsDataTable.Clear()
              }
              $objActivitiesDataTable.Clear()
          }
          $objProcessesDataTable.Clear()
      }
  }
  Process{
    ,$ActList
  }
  End{
    $objOleDbConnection.Close()
    
    $RootLog.Info($MsgTable.StopMsg)
    
    Stop-LoggerSvc
    
    $ModList | %{Remove-Module $_}
  }
}

. Main
# SIG # Begin signature block
# MIIEMwYJKoZIhvcNAQcCoIIEJDCCBCACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUN259/F2DVCMU5YAVEHXq60wo
# sQygggI9MIICOTCCAaagAwIBAgIQtBirZz3Acb1BfUstCv49PTAJBgUrDgMCHQUA
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
# FJ4pbCQRnVDCjWb57z9W2nH0IXcEMA0GCSqGSIb3DQEBAQUABIGAWjKjtPhuQZqy
# vcC5mAyqs/TUB1aXn41Eqbm05shud+fvvQVjfo9DSWqd03h9FyQddFXhxHmNWUGi
# 1JaevBy57Apk0n+uge5/l9e/O4OcjiZLtDzNXogQXXagsF1VCCyeTxa1x0fRD6nM
# SyRgp9hrlXepIKZ1Z6O6T07ovtkp2pI=
# SIG # End signature block
