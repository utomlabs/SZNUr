Param(
	[Parameter(Mandatory = $true, Position = 0)]
	[Alias("Dll")]
	[ValidateNotNullOrEmpty()]
	[string]
	# Log4net dll path
	$log4netDllPath
)

Function Start-LoggerSvc {
<#
.SYNOPSIS
      This function creates a log4net logger instance already configured
.OUTPUTS
      The log4net logger "root" instance ready to be used
#>
[CmdletBinding()]
Param(
	[Parameter(Mandatory = $true, Position = 0)]
	[string]
	# Path of the configuration file of log4net
	$Configuration
)
	Set-StrictMode -Version Latest
	Write-Verbose "[Start-LoggerSvc] Logger initialization"
	$log4netDllPath = Resolve-Path $log4netDllPath -ErrorAction SilentlyContinue -ErrorVariable Err
	if ($Err){
		throw "Log4net library cannot be found on the path"
	}else{
		Write-Verbose "[Start-LoggerSvc] Log4net dll path is : '$log4netDllPath'"
		[void][Reflection.Assembly]::LoadFrom($log4netDllPath) | Out-Null
		# Log4net configuration loading
		$log4netConfigFilePath = Resolve-Path $Configuration -ErrorAction SilentlyContinue -ErrorVariable Err
		if ($Err){
			throw "Log4Net configuration file cannot be found"
        }else {
			Write-Verbose "[Start-LoggerSvc] Log4net configuration file is '$log4netConfigFilePath' "
			$FileInfo = New-Object System.IO.FileInfo($log4netConfigFilePath)
			[log4net.Config.XmlConfigurator]::Configure($FileInfo)
			$SCRIPT:MyCommonLogger = [log4net.LogManager]::GetLogger("root")
			Write-Verbose "[Start-LoggerSvc] Logger service is configured"
			return $MyCommonLogger
        }
    }
}
    
Function Get-Logger {
<#
.SYNOPSIS
      This function returns log4net logger already initiated instance
.OUTPUTS
      The log4net logger instance
.NOTES
      Loggers should be created with log4net configuration file. This function finds and returns already initiated logger instance. It does not create new instance.
#>
[CmdletBinding()]
Param(
	[Alias("ln")]
	[string] $LoggerName
)
	Set-StrictMode -Version Latest
	Write-Verbose "[Get-Logger] Checking if 'log4net' assembly is already loaded"

	if(([appdomain]::currentdomain.getassemblies() | Where {$_ -match "log4net"}) -eq $null){
		throw "log4net tool library is not initializated yet"
	}
    else {
        Write-Verbose "[Get-Logger] log4net tool library is already initializated"
        if([log4net.LogManager]::Exists("root") -eq $null){
            throw "Even 'root' logger is not initializated yet"
        }else {
            Write-Verbose "[Get-Logger] 'root' logger is already initializated"
            $script:MyLogger = [log4net.LogManager]::Exists($LoggerName)
            if($MyLogger -eq $null){
                throw "Logger $LoggerName is not configured yet"
            }else {
                Write-Verbose "[Get-Logger] Logger $LoggerName is configured"
                return $MyLogger
            }
        }
    }
}

Function Add-Logger{
<#
.SYNOPSIS
      This function creates subsequent log4net logger instances
.OUTPUTS
      The log4net logger instance ready to be used
.NOTES
      It is not recommended to use Add-Logger function. Loggers should be created with log4net configuration file.
#>
[CmdletBinding()]
Param(
	[Alias("ln")]
	[string] $LoggerName
)
	Set-StrictMode -Version Latest
	Write-Verbose "[Add-Logger] Checking if 'log4net' assembly is already loaded"

	if(([appdomain]::currentdomain.getassemblies() | Where {$_ -match "log4net"}) -eq $null){
		throw "log4net tool library is not initializated yet"
    }else {
		Write-Verbose "[Add-Logger] log4net tool library is already initializated"
        if([log4net.LogManager]::Exists("root") -eq $null){
            throw "Even 'root' logger is not initializated yet"
        }else {
            Write-Verbose "[Add-Logger] 'root' logger is already initializated"
            $script:MyLogger = [log4net.LogManager]::GetLogger($LoggerName)
            Write-Verbose "[Add-Logger] Logger is configured"
            return $MyLogger
        }
    }
}

Function Stop-LoggerSvc{
<#
.SYNOPSIS
      This function closes log4net session
.OUTPUTS
      None
.NOTES
      None
#>
[CmdletBinding()]
Param(
)
	Set-StrictMode -Version Latest
	Write-Verbose "[Stop-LoggerSvc] Checking if 'log4net' assembly is already loaded"

	if(([appdomain]::currentdomain.getassemblies() | Where {$_ -match "log4net"}) -eq $null){
        throw "log4net tool library is not initializated yet"
    }else {
        Write-Verbose "[Stop-LoggerSvc] log4net tool library is already initializated"
        [log4net.LogManager]::Shutdown()
        Write-Verbose "[Stop-LoggerSvc] Logger service is closed"
    }
}

Export-ModuleMember -Function @(
'Start-LoggerSvc',
'Get-Logger',
'Add-Logger',
'Stop-LoggerSvc'
)

# SIG # Begin signature block
# MIIEMwYJKoZIhvcNAQcCoIIEJDCCBCACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZjw1BBXEJGETHZZp0GOe0Os1
# xO2gggI9MIICOTCCAaagAwIBAgIQtBirZz3Acb1BfUstCv49PTAJBgUrDgMCHQUA
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
# FMk8wGFKXAgsUOrzkazkbUdfVrJyMA0GCSqGSIb3DQEBAQUABIGAicpW1gVxEAEe
# vV+bES6+awYVYSUCtL5DayGcaVpWkytWHXXs3TWhHKoJpw8w9CnCAcQPgfL6juBd
# BHfOWn0xfxlZT3V+H4zfYbZs9Y2COz2MqFbkC23HZhaBIQsP1Ri/6gvXwgyykwy2
# 2m80X+DOJQ0Jmvgy+BWN3bZzLzaVupw=
# SIG # End signature block
