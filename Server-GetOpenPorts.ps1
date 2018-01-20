#Returns open ports on local computer
#Params filter port number, process name/description
#Run on remote computers, import list, export to excel

function Server-GetOpenPorts {
    param (
        [switch]$GetHelp,
	    [array]$ComputerName,
	    [string]$ExportExcel,
	    [string]$ImportFromText,
        [array]$Port,
        [array]$Process,
        [switch]$ExtendedProperties,
        [int]$ThreadCount = ((Get-WmiObject -Class Win32_Processor | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum * 10)
    )
    
    function getPorts { #Main function to retieve open ports
    
        $processfilter = @();

        if (!$Process) {
            $processfilter = "*";
        }

        foreach ($p in $Process) {
        
            $processfilter += $p;

        }

        if ($com -eq $env:COMPUTERNAME) { 
            $ports = netstat -ano | findstr "LISTENING";
            $processes = Get-Process $processfilter;
            
            if ($ExtendedProperties) {
                localAdminCheck;
            }

            if ($localcheck -eq 1) { #Popup if failed localAdminCheck
                Write-Warning "Process descriptions and paths may not be shown for local computer $env:COMPUTERNAME.`n         Please run powershell in administrator mode to view them.";
                #$message = [System.Windows.MessageBox]::Show("Process descriptions and paths may not be shown on local computer $env:COMPUTERNAME.`nPlease run powershell in administrator mode to view them.");
            }

        }

        else {

           $ports = Invoke-Command -ComputerName $com {netstat -ano | findstr "LISTENING"};
       
           if($ports) {
                $processes = Invoke-Command -ComputerName $com {Get-Process $args} -ArgumentList $processfilter;
           }
           else {
                return;
           }

        }

        $tempports = @();
        foreach ($singleport in $Port) {
        
            $tempports += $ports | findstr /R /C:":$singleport ";

        }

        if ($Port) {
            $ports = $tempports;
        }

        if (!($ports -And $processes)) {
            $com = $com.ToUpper()
            Write-Host;
            Write-Host "No results for $com.";
            if ($com -eq $env:COMPUTERNAME) {
                $Script:localcheck = 0;
            }
            return;
        }

        $portdetails = @();

        foreach ($openport in $ports) {
        
            $openport = $openport.substring(2);
            $line = $openport -split "\s+";

            $openprocess = $processes | ?{$_.ID -eq $line[4]}

            if ($openprocess) {           

                $regex = $line[1] | Select-String "(.*):([0-9]+)$";

                $portdetails += New-Object -Type PSCustomObject -Property @{
                    Computer    = $com.ToUpper()
                    Protocol    = $line[0]
                    Interface   = $regex.Matches.Groups[1].Value
                    Port        = $regex.Matches.Groups[2].Value
                    PID         = $openprocess.ID
                    Process     = $openprocess.ProcessName
                    Description = $openprocess.Description
                    Path        = $openprocess.Path
                }

            }

        }
    
        return $portdetails;
    }

    $getportsblock = { #Main function to retieve open ports
        param(
            [array]$Process,
            [array]$Port,
            $com
        )

        $processfilter = @();

        if (!$Process) {
            $processfilter = "*";
        }

        foreach ($p in $Process) {
        
            $processfilter += $p;

        }

        if ($com -eq $env:COMPUTERNAME) { 
            $ports = netstat -ano | findstr "LISTENING";
            $processes = Get-Process $processfilter;                     
        }

        else {

           $ports = Invoke-Command -ComputerName $com {netstat -ano | findstr "LISTENING"};
       
           if($ports) {
                $processes = Invoke-Command -ComputerName $com {Get-Process $args} -ArgumentList $processfilter;
           }
           else {
                return;
           }

        }

        $tempports = @();
        foreach ($singleport in $Port) {
        
            $tempports += $ports | findstr /R /C:":$singleport ";

        }

        if ($Port) {
            $ports = $tempports;
        }

        if (!($ports -And $processes)) {       
            return $null;
        }

        $portdetails = @();

        foreach ($openport in $ports) {
        
            $openport = $openport.substring(2);
            $line = $openport -split "\s+";

            $openprocess = $processes | ?{$_.ID -eq $line[4]}

            if ($openprocess) {           

                $regex = $line[1] | Select-String "(.*):([0-9]+)$";

                $portdetails += New-Object -Type PSCustomObject -Property @{
                    Computer    = $com.ToUpper()
                    Protocol    = $line[0]
                    Interface   = $regex.Matches.Groups[1].Value
                    Port        = $regex.Matches.Groups[2].Value
                    PID         = $openprocess.ID
                    Process     = $openprocess.ProcessName
                    Description = $openprocess.Description
                    Path        = $openprocess.Path
                }

            }

        }
    
        return $portdetails;
    }

    function display {
    
        param (
            $Details
        )

        if ($ExtendedProperties) {
            #$details = $details | Format-Table -Property Protocol, Port, Interface, PID, Process, Description, Path -GroupBy Computer;
            if ($ExportExcel) {
                $Details | Select Computer, Protocol, Port, Interface, PID, Process, Description, Path | Export-Excel $ExportExcel -WorkSheetname "Open Ports";
            }

            else {
                $Details | Format-Table -Property Protocol, Port, Interface, PID, Process, Description, Path -GroupBy Computer;
            }

        }

        else {
            if ($ExportExcel) {
                $Details | Select Computer, Protocol, Port, Interface, Process | Export-Excel $ExportExcel -WorkSheetname "Open Ports";
            }

            else {
                $Details | Format-Table -Property Protocol, Port, Interface, Process -GroupBy Computer;
            }
        }
    }

    function standardParam {
    
        param(
            $Pipe
        )
    
        if (helpmenu -eq 1) {
            return 1;
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

        foreach ($com in $Pipe) {
        
            if ($com.GetType().Name -eq "ADComputer") {

                [array]$script:Computers += $com.Name;
            
            }

	        else {
                [array]$script:Computers += $com;
            }
	
        }

        if ($ImportFromText) {
	        [array]$script:Computers += Get-Content $ImportFromText;
	        if (!$script:Computers) {
		        throw "Invalid input file";
	        }
        }

        if (!$script:Computers) {
        
            $script:Computers = $env:COMPUTERNAME;

        }

    }

    function helpmenu {

        if ($GetHelp -Or ($ComputerName -And $ImportFromText)) {	    
	        Write-Host;
	        Write-Host "Returns open ports on local computer";
	        Write-Host;
	        Write-Host "Usage: $ScriptName [-ComputerName <string[,]> | -ImportFromText <filepath>] [-ExtendedProperties] [-Port <num[,]>] [-Process <string[,]>] [-ExportExcel <filepath>]";		
	        Write-Host;		
	        Write-Host 'Options:';		
	        Write-Host '  -ComputerName <string[,]>    Run the command on remote computer. Seperate multiple names by ","';		
	        Write-Host '  -ImportFromText <filepath>   Import computer names from .txt file';
            Write-Host '  -ExtendedProperties          Shows more detailed properties';
            Write-Host '  -Port <num[,]>               Filter results by port number(s) specified. Seperate mutiple port numbers by ","';
            Write-Host '  -Process <string[,]>         Filter results by process name(s) specified. Accepts wildcards. Seperate multiple names by ","';
	        Write-Host '  -ExportExcel <filepath>      Specify a filepath ending in .xlsx to export list, e.g. ".\Downloads\list.xlsx\" ';		
	        Write-Host;
	        return 1;
        }

    }

    function localAdminCheck {
    
        Add-Type -AssemblyName PresentationFramework;
        
        if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
            
            Write-Warning "Process descriptions and paths may not be shown for local computer $env:COMPUTERNAME.`n         Please run powershell in administrator mode to view them.";            
        
        }

    }

    $script:Computers = $ComputerName;
    $ScriptName = $MyInvocation.MyCommand.Name
    if(standardParam -Pipe $input -eq 1) {
        return;
    }

    $progress = 0;
    Write-Progress -Activity "Retrieving open port details" -Status "Completed 0 of $($Computers.Count) computers" -PercentComplete $progress;

    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $ThreadCount);
    $RunspacePool.Open();
    $RunningJobs = New-Object System.Collections.ArrayList;

    foreach ($com in $script:Computers) {
        
        if ($com -eq $env:COMPUTERNAME -and $ExtendedProperties) {            
            localAdminCheck;
        }

        [System.Management.Automation.PowerShell]$PSThread = [System.Management.Automation.PowerShell]::Create();
        $PSThread.RunspacePool = $RunspacePool;
        [void]$PSThread.AddScript($getportsblock);
        [void]$PSThread.AddParameter("com", $com).AddParameter("Process", $Process).AddParameter("Port", $Port);      
        $Handle = $PSThread.BeginInvoke();
        
        $ThreadObj = '' | Select Thread, Handle, Computer;
        $ThreadObj.Thread = $PSThread;
        $ThreadObj.Handle = $Handle;
        $ThreadObj.Computer = $com;
        

        [void]$RunningJobs.Add($ThreadObj);

    }

    while($RunningJobs.Count -gt 0) {
        
        Start-Sleep -Seconds 5;

        $templist = $RunningJobs.Clone();

        $RunningJobs | % {
            
            $Handle = $_.Handle;

            if ($Handle.IsCompleted) {
                
                $list = $_.Thread.EndInvoke($Handle);
                [void]$_.Thread.Dispose();

                if ($list) {
                    $details += $list;
                }
                else {
                    Write-Output "No results found for $($_.Computer.ToUpper()).";
                }
                
                $progress++;
                Write-Progress -Activity "Retrieving open port details" -Status "Completed $progress of $($Computers.Count) computers" -PercentComplete (($progress/$Computers.Length) * 100);
                [void]$templist.Remove($_);

            }

        }

        $RunningJobs = $templist;    

    }
    
    $RunspacePool.Close();
    $RunspacePool.Dispose();

    display -Details $details;

}

# SIG # Begin signature block
# MIIMIQYJKoZIhvcNAQcCoIIMEjCCDA4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUXiZfiAaGFlD/wLpn2eEqLQUP
# vy+gggoMMIIE7DCCBFWgAwIBAgIKE7/jkQAAAAAADjANBgkqhkiG9w0BAQUFADBS
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
# BgkqhkiG9w0BCQQxFgQUNsMufioErjIZX54x+PkFdbZgFhYwDQYJKoZIhvcNAQEB
# BQAEgYB847HHOek2i3izOmSfHqx/+E7Ybx/fu4cElXFooegkBlkFS2FSEyWNeHaF
# ul2av27f6HE4/rmKr5NaYbk5NpBNMc5ilgJtQvVeGVcSKcRElPTC0QmI5eM/Xsje
# 0mntMRF/LzCL7DuTi/dI3CZCt77s0479R2J4wMdAlLHSNPqdAQ==
# SIG # End signature block
