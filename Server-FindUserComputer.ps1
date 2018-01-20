#  Find the computer which a user was using

#  Input ADUser, Output Computers where ManagedBy = ADUser
function Server-FindUserComputer {
    param (
        [array]$User,
        [ValidateSet('Admin Accounts', 'CEO Office', 'Disabled Users', 'Finance', `
                     'HR', 'Sales')]
        [string]$SearchBase,
        [string]$ExportExcel,
	    [string]$ImportFromText,
	    [switch]$GetHelp
    )

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

        foreach ($usr in $Pipe) {
        
            if ($usr.GetType().Name -eq "ADUser") {

                [array]$script:ADUsers += $usr.SamAccountName;
            
            }

	        else {
                [array]$script:ADUsers += $usr;
            }
	
        }

        if ($ImportFromText) {
            [array]$script:ADUsers += Get-Content $ImportFromText;
	        if (!$script:ADUsers) {
		        throw "Invalid input file";
	        }
        }

    }

    function helpmenu {

        if ($GetHelp -Or ($User -And $ImportFromText)) {
	        Write-Host;
	        Write-Host "Returns list of computers used by user(s)";
	        Write-Host;
	        Write-Host "Usage: $ScriptName [-User <user[,]> | -ImportFromText <filepath>] [-SearchBase <base>] [-ExportExcel <filepath>]";		
	        Write-Host;
            Write-Host 'Options:';		
	        Write-Host '  -User <user[,]>              Specify users to check. Seperate multiple names by ","';
            Write-Host '  -ImportFromText <filepath>   Import users from .txt file';
            Write-Host '  -SearchBase <base>           Specify department to search for users from';
            Write-Host '  -ExportExcel <filepath>      Specify a filepath ending in .xlsx to export list, e.g. ".\Downloads\list.xlsx\" ';	
            Write-Host;	
	        return 1;
        }

    }

    function getComputer {

        $sam = $ManagedBy.SamAccountName;
        
        $com = '';
        $com = Get-ADComputer -Server $ADServer -Filter {ManagedBy -eq $sam -And Enabled -eq "True"} `
                        -Properties Name, CanonicalName, ManagedBy, Modified, IPv4Address, `
                                    OperatingSystem, OperatingSystemServicePack | `
                        Select @{name='Computer Name';expression={$_.Name}}, `
                                @{name='Location';expression={$_.CanonicalName}}, `
                                @{name='Last Used On';expression={$_.Modified}}, `
                                @{name='IP Address';expression={$_.IPv4Address}}, `
                                @{name='Operating System';expression={$_.OperatingSystem + ' ' + $_.OperatingSystemServicePack}};

        foreach ($c in $com) {	
    
            $c.Location = $c.Location -replace '([.]*)/[0-9a-zA-Z\.\- ]*$', $1;		

        }

        $name = $sam;
        $location = $ManagedBy.CanonicalName -replace '([.]*)/[0-9a-zA-Z\.\- ]*$', $1;

        if (!$ExportExcel) {

            if ($com) {
                Write-Host "Name     : $($ManagedBy.Name)";
                Write-Host "ID       : $name";
                Write-Host "Location : $location";
                $com | Format-Table;

            }

            else {
                Write-Host "Name     : $($ManagedBy.Name)";
                Write-Host "ID       : $name";
                Write-Host "Location : $location";
                Write-Host;
                Write-Host "No computers found for the user.";
                Write-Host;
            }    

        }

        else {

            return $com | Select @{name='User Name';expression={$ManagedBy.Name}}, `
                                 @{name='User ID';expression={$name}}, `
                                 @{name='User Location';expression={$location}}, `
                                 'Computer Name', `
                                 @{name='Computer Location';expression={$_.Location}}, `
                                 'Last Used On', `
                                 'IP Address', `
                                 'Operating System'

        }

    }

    $script:ADUsers = $User;
    $ScriptName = $MyInvocation.MyCommand.Name;
    if(standardParam -Pipe $input -eq 1) {
        return;
    }
    Add-Type -AssemblyName PresentationFramework;

    Write-Host;
    
    switch($SearchBase) {
        
        'Admin Accounts'  {$sb = 'OU=Admin Accounts,OU=USER,DC=EXAMPLE,DC=COM,DC=SG';}
        'CEO Office'      {$sb = 'OU=CEO Office,OU=USER,DC=EXAMPLE,DC=COM,DC=SG';}
        'Disabled Users'  {$sb = 'OU=Disabled Users,OU=USER,DC=EXAMPLE,DC=COM,DC=SG';}
        'Finance'         {$sb = 'OU=Finance,OU=USER,DC=EXAMPLE,DC=COM,DC=SG';}
        'HR'              {$sb = 'OU=HR,OU=USER,DC=EXAMPLE,DC=COM,DC=SG';}
        'Sales'           {$sb = 'OU=Sales,OU=USER,DC=EXAMPLE,DC=COM,DC=SG';}
        default           {$sb = 'DC=EXAMPLE,DC=COM,DC=SG';}

    }

    $details = @();

    if (!$script:ADUsers) {
        
        $script:ADUsers = Get-ADUser -Server $ADServer -SearchBase $sb -Filter * -Properties CanonicalName, SamAccountName, EmailAddress | Select Name, SamAccountName, EmailAddress, CanonicalName | Out-GridView -Title "Select user" -PassThru;
        
        foreach ($u in $script:ADUsers) {

            $ManagedBy = $u;
            
            if ($ExportExcel) {
                $details += getComputer;
            }
            
            else {
                getComputer;
            }

        }

    }

    else {
        
        foreach ($u in $script:ADUsers) {
            $Identity = '';
	        $Identity = Get-ADUser -Server $ADServer -SearchBase $sb -Filter {Name -Like $u -Or SamAccountName -Like $u} -Properties CanonicalName, SamAccountName, EmailAddress;
        
            if (!$Identity) {
                Write-Host "Unable to find any user matching '$u'";
                Write-Host;
                continue;
            }

            if ($Identity.Count -gt 1) {
                $message = [System.Windows.MessageBox]::Show("There are multiple matches for '$u'. Please select the correct user.");
                $ManagedUser = $Identity | Select Name, SamAccountName, EmailAddress, CanonicalName | Out-GridView -Title "Select user matching $u" -PassThru;
            }
            else {
                $ManagedUser = $Identity | Select Name, SamAccountName, EmailAddress, CanonicalName;
            }
        
            if (!$ManagedUser) {
                return;
            }

            if($ManagedUser.Count -gt 1){ 
                
                foreach ($usr in $ManagedUser) {
                    
                    $ManagedBy = $usr;
                    
                    if ($ExportExcel) {
                        $details += getComputer;
                    }
                    
                    else {
                        getComputer;
                    }

                }

            }
            
            else {
                $ManagedBy = $ManagedUser;
                
                if ($ExportExcel) {
                    $details += getComputer;
                }

                else {
                    getComputer;
                }

            }

        }

    }

    if ($details) {
        $details | Export-Excel $ExportExcel -WorkSheetname "Computers Used By";
    }

}

# SIG # Begin signature block
# MIIMIQYJKoZIhvcNAQcCoIIMEjCCDA4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUeGAnlhrbuP1AeJUEYtHIRk8y
# bFygggoMMIIE7DCCBFWgAwIBAgIKE7/jkQAAAAAADjANBgkqhkiG9w0BAQUFADBS
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
# BgkqhkiG9w0BCQQxFgQUYaAwI+grOjFmjRaoORKOZgYCzpcwDQYJKoZIhvcNAQEB
# BQAEgYBWYLOv3lFlSqWRYhqrRaQGCaKThvGx8a3caKRNlwSgVUOaVREzWZyY5a8K
# H/IUMm4ErEdreW3VC+RIIYEqS17wu+VqPjCVYWuqZqtqqRz5aGFDySsKxCNJYQAG
# JY9dsJ8cqsvx9AWedSIeeFGpmLl0E2AAUjE4nRWrt9i/r3MhKA==
# SIG # End signature block
