function Server-GetInstalledPrograms {
    param (
	    [array]$ProgramName,
        [switch]$GetHelp,
	    [array]$ComputerName,
	    [string]$ExportExcel,
	    [string]$ImportFromText,
        [int]$ThreadCount = ((Get-WmiObject -Class Win32_Processor | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum * 10)
    )

    $remoteprogramsblock = {
        param (
            $Computer,
            $ToFilter
        )

        try {
            #open registry path containing installed info
            $w32reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer);
            $keypath = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\';
            $reg = $w32reg.OpenSubKey($keypath);

            $list = @();
        
            foreach ($key in $reg.GetSubKeyNames()) { #opens each folder in path and extract program info

                $programreg = $reg.OpenSubKey($key);

                if($programreg.GetValue("DisplayName")) {

                    $obj = New-Object -Type PSCustomObject -Property @{
                        Computer       = $Computer.ToUpper()
                        "Program Name" = $programreg.GetValue("DisplayName")
                        Version        = $programreg.GetValue("DisplayVersion")
                        Publisher      = $programreg.GetValue("Publisher")
                        "Installed On" = $programreg.GetValue("InstallDate") -replace "^([0-9]{4})([0-9]{2})([0-9]{2})$", '$3/$2/$1'
            
                    }

                    $list += $obj;
                }

            }
                      
            if ($ToFilter) { #filter list by program name
            
                $filter = @();

                foreach ($p in $ToFilter) {
         
                    $filter += $list | ? {$_."Program Name" -like "$p"}

                }

                $list = $filter;

            }
            
            return $list;

        }
        catch {
            #catches error while connecting
            Write-Error "Failed to connect to $Computer";           

        }

    }

    if ($GetHelp -Or ($ComputerName -And $ImportFromText)) { #helpmenu
	    $ScriptName = $MyInvocation.MyCommand.Name
	    Write-Host;
	    Write-Host "Returns programs installed on local computer";
	    Write-Host;
	    Write-Host "Usage: $ScriptName [-ProgramName <string[,]>] [-ComputerName <string[,]> | -ImportFromText <filepath>] [-ExportExcel <filepath>]";		
	    Write-Host;		
	    Write-Host 'Options:';
        Write-Host '  -ProgramName <string[,]>     Filters results to program(s) specified. Accepts wildcards. Seperate multiple programs by ","';		
	    Write-Host '  -ComputerName <string[,]>    Run the command on remote computer. Seperate multiple names by ","';		
	    Write-Host '  -ImportFromText <filepath>   Import computer names from .txt file';
	    Write-Host '  -ExportExcel <filepath>      Specify a filepath ending in .xlsx to export list, e.g. ".\Downloads\list.xlsx\" ';	
        Write-Host '  -ThreadCount <int>           Specify the number of threads to run concurrently. Default is No. of CPUs * 10. Higher number = Faster Processing but Larger CPU/Memory Usage';		
	    Write-Host;
	    return;
    }

    if ($ExportExcel) {
            
        if (!($ExportExcel -Like "*.xlsx")) {
	        throw "Filepath must end in '.xlsx'";
        }

        [ref]$ExportPath = [ref]$ExportExcel;
    }
            
    if ($ExportPath) {
            
        if ($ExportPath.Value -like "\*") {
            $ExportPath.Value = $ExportPath.Value.Substring(1);
        }

        if ($ExportPath.Value -like ".\*") {
            $ExportPath.Value = (pwd).Path + $ExportPath.Value.Substring(1);
        }

        elseif (!($ExportPath.Value -like "[a-z]:\*")) {
            $ExportPath.Value = (pwd).Path + "\" + $ExportPath.Value;
        }

        $ExportPath.Value -match "(.*\\).*\.[a-z]{4}$" | Out-Null
        $TestPath = $Matches[1];

        if (!(Test-Path $TestPath)) {
            throw "Invalid path $($ExportPath.Value)";
        }
    }

    $programlist = @();

    foreach ($com in $input) { #pipeline handling

	    if ($com.GetType().Name -eq "ADComputer") { #accept ADComputer objects in pipeline

            $ComputerName += $com.Name;
            
        }

        else {
            $ComputerName += $com;
	    }

    }

    if ($ImportFromText) { #extract computers from text file
	    $ComputerName += Get-Content $ImportFromText;
	    if (!$ComputerName) {
		    throw "Invalid input file";
	    }
    }

    if (!$ComputerName) { #default computer
        $ComputerName = $env:COMPUTERNAME;
    }
    
    $progress = 0;
    Write-Progress -Activity "Retrieving Installed Programs" -Status "Completed 0 of $($ComputerName.Count) computers" -PercentComplete $progress;

    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $ThreadCount);
    $RunspacePool.Open();
    $RunningJobs = New-Object System.Collections.ArrayList;

    foreach ($c in $ComputerName) {
        
        [System.Management.Automation.PowerShell]$PSThread = [System.Management.Automation.PowerShell]::Create();
        $PSThread.RunspacePool = $RunspacePool;
        [void]$PSThread.AddScript($remoteprogramsblock);
        [void]$PSThread.AddParameter("Computer", $c).AddParameter("ToFilter", $ProgramName);      
        $Handle = $PSThread.BeginInvoke();
        
        $ThreadObj = '' | Select Thread, Handle, Computer;
        $ThreadObj.Thread = $PSThread;
        $ThreadObj.Handle = $Handle;
        $ThreadObj.Computer = $c;
        

        [void]$RunningJobs.Add($ThreadObj);

    }

    while($RunningJobs.Count -gt 0) {
        
        Start-Sleep -Seconds 5;

        $templist = $RunningJobs.Clone();

        $RunningJobs | % {
            
            $Handle = $_.Handle;

            if ($Handle.IsCompleted) {
                
                $proglist = $_.Thread.EndInvoke($Handle);
                [void]$_.Thread.Dispose();

                if ($proglist) {
                    $programlist += $proglist;
                }
                else {
                    Write-Output "No results found for $($_.Computer.ToUpper()).";
                }
                
                $progress++;
                Write-Progress -Activity "Retrieving Installed Programs" -Status "Completed $progress of $($ComputerName.Count) computers" -PercentComplete (($progress/$ComputerName.Length) * 100);
                [void]$templist.Remove($_);

            }

        }

        $RunningJobs = $templist;    

    }

    $RunspacePool.Close();
    $RunspacePool.Dispose();
    

    if($ExportExcel) { #export list to excel
	
	    $programlist | Select -Property "Computer", "Program Name", "Installed On", "Version", "Publisher" | Export-Excel $ExportExcel -WorkSheetname "Installed Programs";

    }

    else {
        $programlist | Format-Table -AutoSize -GroupBy "Computer" -Property "Program Name", "Installed On", "Version", "Publisher";
    }

}

# SIG # Begin signature block
# MIIMIQYJKoZIhvcNAQcCoIIMEjCCDA4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUEZA/fsq/17HjtJK5Ct9WyZHI
# WhegggoMMIIE7DCCBFWgAwIBAgIKE7/jkQAAAAAADjANBgkqhkiG9w0BAQUFADBS
# MRIwEAYKCZImiZPyLGQBGRYCU0cxEzARBgoJkiaJk/IsZAEZFgNDT00xEzARBgoJ
# kiaJk/IsZAEZFgNNSVQxEjAQBgNVBAMTCU1JVFJPT1RDQTAeFw0xNzAyMjcwMzE0
# NDBaFw0xOTAyMjcwMzI0NDBaMFExEjAQBgoJkiaJk/IsZAEZFgJTRzETMBEGCgmS
# JomT8ixkARkWA0NPTTETMBEGCgmSJomT8ixkARkWA01JVDERMA8GA1UEAxMITUlU
# U1VCQ0EwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBALI0XZIoLMinWfRZfj5N
# SyumTWTYuHnQVFQlVpwDdvl+a7WKgJh0jXIJDuvr4Uy/+kOcaQsV8vCr5S3I2gtU
# Rz/8/eZF4kJTMilOYTUd7Ld1KT6yOODAk4BrTBfuXEss4ZJ/ulSyuYypriVz/sTX
# d7iy/MQeRvP/9bX75jU5yzFDAgMBAAGjggLIMIICxDAPBgNVHRMBAf8EBTADAQH/
# MB0GA1UdDgQWBBRB2t93nYIXT3jbqEc81V90/Ber0jALBgNVHQ8EBAMCAYYwEAYJ
# KwYBBAGCNxUBBAMCAQAwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0j
# BBgwFoAUD2P8MIHgWU9Tl1zJlmKqbzNbKVQwggEPBgNVHR8EggEGMIIBAjCB/6CB
# /KCB+YaBumxkYXA6Ly8vQ049TUlUUk9PVENBLENOPWV4Y2htYWlsLXNlcnZlcixD
# Tj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049
# Q29uZmlndXJhdGlvbixEQz1NSVQsREM9Q09NLERDPVNHP2NlcnRpZmljYXRlUmV2
# b2NhdGlvbkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2lu
# dIY6aHR0cDovL2V4Y2htYWlsLXNlcnZlci5taXQuY29tLnNnL0NlcnRFbnJvbGwv
# TUlUUk9PVENBLmNybDCCASIGCCsGAQUFBwEBBIIBFDCCARAwgaoGCCsGAQUFBzAC
# hoGdbGRhcDovLy9DTj1NSVRST09UQ0EsQ049QUlBLENOPVB1YmxpYyUyMEtleSUy
# MFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9TUlULERD
# PUNPTSxEQz1TRz9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlm
# aWNhdGlvbkF1dGhvcml0eTBhBggrBgEFBQcwAoZVaHR0cDovL2V4Y2htYWlsLXNl
# cnZlci5taXQuY29tLnNnL0NlcnRFbnJvbGwvZXhjaG1haWwtc2VydmVyLk1JVC5D
# T00uU0dfTUlUUk9PVENBLmNydDANBgkqhkiG9w0BAQUFAAOBgQDAMjQJaLVRbZoh
# s3MTUZRV3gxpZcvo2w87TapB4ZQVaRCY9Uqsg3yswOaUyvQ3ZMXGo6l4C7ccf/WW
# ZJpoq9YsDkpSimGN4HIm5HZir71gB5/OWVgkgyiMtKLbkesykNewvfb9rHJLjOlm
# VbMOVDzNox3cT+6Chl2cByjjNaHJbDCCBRgwggSBoAMCAQICCl5KrUkAAAAAACAw
# DQYJKoZIhvcNAQEFBQAwUTESMBAGCgmSJomT8ixkARkWAlNHMRMwEQYKCZImiZPy
# LGQBGRYDQ09NMRMwEQYKCZImiZPyLGQBGRYDTUlUMREwDwYDVQQDEwhNSVRTVUJD
# QTAeFw0xNzA1MDUwNTU0MzBaFw0xODA1MDUwNTU0MzBaMHgxEjAQBgoJkiaJk/Is
# ZAEZFgJTRzETMBEGCgmSJomT8ixkARkWA0NPTTETMBEGCgmSJomT8ixkARkWA01J
# VDEQMA4GA1UECxMHTUlUVVNFUjEMMAoGA1UECxMDTUlTMRgwFgYDVQQDEw9DaGVv
# bmcgUmVuIEhhbm4wgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBAOUhRgYlnyoS
# e3l8bWbeI7tAnKcifLWSJc5vF+23HkjB57FwgRVlecXW2KoPxWCTlR1PgzDhssSG
# fJHvHsgeHasUzMRGGmCfTALrXjBaU3/028kx6lV3lWgm3KSv43T54wBJQW/0SmN5
# AeyvoiqSXNAyIp7prDB7DA4rt9wYH2bXAgMBAAGjggLOMIICyjAlBgkrBgEEAYI3
# FAIEGB4WAEMAbwBkAGUAUwBpAGcAbgBpAG4AZzATBgNVHSUEDDAKBggrBgEFBQcD
# AzALBgNVHQ8EBAMCB4AwHQYDVR0OBBYEFJeOroVbCXMNA/SoEgOxUtK4hSY9MB8G
# A1UdIwQYMBaAFEHa33edghdPeNuoRzzVX3T8F6vSMIH7BgNVHR8EgfMwgfAwge2g
# geqggeeGgbFsZGFwOi8vL0NOPU1JVFNVQkNBLENOPW1pdGZzMDEsQ049Q0RQLENO
# PVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3Vy
# YXRpb24sREM9TUlULERDPUNPTSxEQz1TRz9jZXJ0aWZpY2F0ZVJldm9jYXRpb25M
# aXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnSGMWh0dHA6
# Ly9taXRmczAxLm1pdC5jb20uc2cvQ2VydEVucm9sbC9NSVRTVUJDQS5jcmwwggEP
# BggrBgEFBQcBAQSCAQEwgf4wgakGCCsGAQUFBzAChoGcbGRhcDovLy9DTj1NSVRT
# VUJDQSxDTj1BSUEsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2Vydmlj
# ZXMsQ049Q29uZmlndXJhdGlvbixEQz1NSVQsREM9Q09NLERDPVNHP2NBQ2VydGlm
# aWNhdGU/YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MFAG
# CCsGAQUFBzAChkRodHRwOi8vbWl0ZnMwMS5taXQuY29tLnNnL0NlcnRFbnJvbGwv
# bWl0ZnMwMS5NSVQuQ09NLlNHX01JVFNVQkNBLmNydDAuBgNVHREEJzAloCMGCisG
# AQQBgjcUAgOgFQwTcmhjaGVvbmdATUlULkNPTS5TRzANBgkqhkiG9w0BAQUFAAOB
# gQAnIdD6IywIhoQXGxWD588rCWukxBxxkGCOEU9+t4ryuATodc4AExUdOSTRA/Ce
# nHKJJqT7WqEM9YFzl7Nahkvu4c98N7pph1uk4mS6QHcjCFUdYNCal2ahyRf7vUu+
# OEOa2DZ3TDSA7nCmoe1mNSLQZeUdDPYedIGuurzs6qYL/DGCAX8wggF7AgEBMF8w
# UTESMBAGCgmSJomT8ixkARkWAlNHMRMwEQYKCZImiZPyLGQBGRYDQ09NMRMwEQYK
# CZImiZPyLGQBGRYDTUlUMREwDwYDVQQDEwhNSVRTVUJDQQIKXkqtSQAAAAAAIDAJ
# BgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAj
# BgkqhkiG9w0BCQQxFgQUGpDmVrRtqOzr0VCAU0EZFvmxNoYwDQYJKoZIhvcNAQEB
# BQAEgYCrsKOMVJRyYvGY0wePX5mJ/Fbdhi909qxBan/ATDeLu/kHRa1P3WF09yrE
# Q14Io5gu0HMRLtoSZhFxVwQqmgfNmyg385VF54sWPUKsk1ZQO65BSqUa8X1eBjZf
# B6Z6d0uH+AEXZO0x8DnEWMHSmJZEbMAXNTlJwQDSricTsr+iAg==
# SIG # End signature block
