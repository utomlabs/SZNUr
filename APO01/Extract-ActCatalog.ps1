## .ExternalHelp C:\Users\IK0212141\Documents\WindowsPowerShell\SZNUr\APO01\Extract-ActCatalog.ps1-help.xml

[CmdletBinding(SupportsShouldProcess=$true)]
Param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="W jakim pliku znajduje się źródłowy regulamin organizacyjny?")]
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
Begin{}
Process{
    Set-StrictMode -Version Latest

    Import-LocalizedData -BindingVariable MsgTable

    Clear-History
    Clear-Host

    try{
        $ModList = New-Object System.Collections.ArrayList
        [void] $ModList.Add(@(Import-Module -Name PSLog -ArgumentList "C:\Users\IK0212141\Documents\WindowsPowerShell\Libs\log4net\bin\net\3.5\release\log4net.dll" -Force -PassThru))
		$RootLog = Start-LoggerSvc -Configuration "C:\Users\IK0212141\Documents\WindowsPowerShell\SZNUr\APO01\Extract-ActCatalog.ps1.config"
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

    $RootLog.Info($MsgTable.StartMsg)

    $ScriptLog = Get-Logger -ln Extract-ActCatalog

    ## As Microsoft.ACE.OLEDB.12.0 provider works only on x86 architecture the following statement acts as hardening of the basic functional code
	$ScriptLog.Info($env:Processor_Architecture)
    if ($env:Processor_Architecture -ne "x86") {
        $ScriptLog.Warn("Running x86 PowerShell...")
        $RootLog.Info($MsgTable.StopMsg)
        Stop-LoggerSvc
        if ($myInvocation.Line){
            &"$env:WINDIR\syswow64\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
        } else {
            &"$env:WINDIR\syswow64\windowspowershell\v1.0\powershell.exe" -NonInteractive -File "$($myInvocation.InvocationName)" $args
        }
        exit $lastExitCode
    }
    
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
				}
				$objDetailsDataTable.Clear()
			}
			$objActivitiesDataTable.Clear()
		}
		$objProcessesDataTable.Clear()
	}

    $objOleDbConnection.Close()
	
    $RootLog.Info($MsgTable.StopMsg)

    Stop-LoggerSvc
	
    $ModList | %{Remove-Module $_}
}
End{}
# SIG # Begin signature block
# MIIEMwYJKoZIhvcNAQcCoIIEJDCCBCACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU7I59EXRinSPTEg4bEuu5BsZv
# muegggI9MIICOTCCAaagAwIBAgIQtBirZz3Acb1BfUstCv49PTAJBgUrDgMCHQUA
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
# FLd4Tc1GBfIGljpg+8U0w8lnEhfGMA0GCSqGSIb3DQEBAQUABIGAdEgyM28xakwu
# 3sX/nAOMxeb+cpIy3ecdA2ci1Fr95Jv+5Hbsqwbk3Uj22rlON5T5kBkVCHLOALKz
# Zj87Dz8AHpykdRecaWeJ2KuACjqDT8Abv7fWfAvFNLyNiP4ZxIzQmdGx0ui1lien
# wicf3w0FVWeQm9IpoZ9loVbCLQdhgGE=
# SIG # End signature block
