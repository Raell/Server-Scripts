<# 
Gets VM details from Hyper-V server

To get:

    VM Name - Get-VM -ComputerName SERVER01
    Computer Name - Use nslookup from IP?

    Collection Name - Get-RDVirtualDesktop -ConnectionBroker SERVER.EXAMPLE.COM | Select CollectionName
    Assigned User - Get-RDPersonalVirtualDesktopAssignment -ConnectionBroker SERVER.EXAMPLE.COM -VirtualDesktopName Store01 -CollectionName Store | Select User

    Memory - Get-VM -ComputerName SERVER01

    State  - Get-VM -ComputerName SERVER01
    Uptime - Get-VM -ComputerName SERVER01

    Virtual Disk Details
        Path           - Get-VM -ComputerName SERVER01 | Select VMId | Get-VHD -ComputerName SERVER01 | Select Path
        Type           - Get-VM -ComputerName SERVER01 | Select VMId | Get-VHD -ComputerName SERVER01 | Select VHDType
        Capacity       - Get-VM -ComputerName SERVER01 | Select VMId | Get-VHD -ComputerName SERVER01 | Select Size
        
        Current Size   - Get-VM -ComputerName SERVER01 | Select VMId | Get-VHD -ComputerName SERVER01 | Select FileSize
        (of checkpoint)
        Effective Size - Get-VM -ComputerName SERVER01 | Select VMId | Get-VHD -ComputerName SERVER01 | Select MinimumSize
        (total size of all checkpoints)

    IP Address - Get-VM -ComputerName SERVER01 | Select NetworkAdapters | Select IPAddresses
    
#>

function Server-GetVMDetails {
    param(
        [switch]$GetHelp,
        [array]$VMName = "*",
        [array]$User = "*",
        [array]$Server = $HyperVHosts,
        [array]$ConnectionBroker = $RDSServers,
        [switch]$VHDProperties,
        [switch]$ExtendedProperties,
        [string]$ExportExcel,
        [string]$ExportWord
    )
    
    function helpmenu {
    	
        if($GetHelp) {    
	        Write-Output;
	        Write-Output "Returns details on VMs in Hyper-V servers";
	        Write-Output;
	        Write-Output "Usage: $ScriptName [-VMName <string[,]>] [-User <string[,]>] [-Server <string[,]>] [-ConnectionBroker <string[,]>] [-VHDProperties] [-ExtendedProperties] [-ExportExcel <filepath> | -ExportWord <filepath>]";		
	        Write-Output;		
	        Write-Output 'Options:';		
	        Write-Output '  -VMName <string[,]>            Return details on VM(s) specified. Both Computer and VM names accepted. '
            Write-Output '                                 Wildcards accepted. Seperate multiple names by ","';	
            Write-Output '  -User <string[,]>              Return details on VM used by specified user(s).'
            Write-Output '                                 Wildcards accepted. Seperate multiple names by ","';	
            Write-Output '  -Server <string[,]>            Specify Hyper-V server(s) to connect to. Seperate multiple servers by ","';
            Write-Output '  -ConnectionBroker <string[,]>  Specify Remote Desktop server(s) to connect to. Seperate multiple servers by ","';
            Write-Output '  -VHDProperties                 Shows VHD details for each VM';
            Write-Output '  -ExtendedProperties            Shows addditional properties for each VM';
	        Write-Output '  -ExportExcel <filepath>        Specify a filepath ending in .xlsx to export list, e.g. ".\Downloads\list.xlsx"';
            Write-Output '  -ExportWord <filepath>         Specify a filepath ending in .docx to export list, e.g. ".\Downloads\list.docx"';		
	        Write-Output;
	        return 1;
        }

    }

    function standardParam {        

        if (helpmenu -eq 1) {
            return 1;
        }    
    
        if ($ExportExcel) {
            
            if (!($ExportExcel -Like "*.xlsx")) {
	            throw "Filepath must end in '.xlsx'";
            }

            [ref]$ExportPath = [ref]$ExportExcel;
        }
            
        if ($ExportWord) {
            
            if (!($ExportWord -Like "*.docx")) {
	            throw "Filepath must end in '.docx'";
            }

            [ref]$ExportPath = [ref]$ExportWord;
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
             
    }  

    function GetVMDetails {
            
        $ProgressCount = 0;

        $VMs = @();
        Write-Progress -Activity "Connecting to Servers" -Status "Retrieving VM list" -PercentComplete $ProgressCount;
		
        $VMs += Get-VM -ComputerName $Server -ErrorAction Stop;
        $ProgressCount += 5;

        $DHCPList = Get-DhcpServerv4Lease -ComputerName $DHCPHyperV -ScopeId $HyperVSubnet | Select IPAddress, HostName;

        $Assignments = @();
        
        Import-Module RemoteDesktop 3>$null;

        $ConnectionBroker |
        % {
            $Conn = $_;       

            if($Conn) {
                if (!($Conn -match "$Domain$")) {
                    $Conn += ".$Domain";
                }
            }
          
            $Collections = Get-RDVirtualDesktop -ConnectionBroker $Conn | ? {$_.CollectionName} | Select CollectionName -Unique;

            $Collections | % {

                $Coll = $_.CollectionName;
                Write-Progress -Activity "Connecting to $Conn" -Status "Retrieving Collection: $Coll" -PercentComplete $ProgressCount;
                $VDA = Get-RDPersonalVirtualDesktopAssignment -ConnectionBroker $Conn -CollectionName $Coll;
                $VDA | % {                  
                    
                    $Assignments += New-Object -Type PSCustomObject -Property @{
                        VMName     = $_.VirtualDesktopName
                        Collection = $Coll
                        User       = $_.User
                    }
                }

                $ProgressCount += ((30/$ConnectionBroker.Count)/$Collections.Count);

            }
        }

        $VMs |
        % {                     
            
            $Name = $_.Name;
            $HostServer = $_.ComputerName;
            Write-Progress -Activity "Retrieving Hyper-V VM Details" -Status "Retrieving Details: $Name" -PercentComplete $ProgressCount;

            $IPAddress = $null;
            if ($_.NetworkAdapters) {
                if($_.NetworkAdapters.IPAddresses.Count -gt 1) {
                    $IPAddress = $_.NetworkAdapters.IPAddresses[0];
                }
                else {
                    $IPAddress = $_.NetworkAdapters.IPAddresses;
                }
            } #Gets IP Address

            $VHD = $VHD = $_.HardDrives.Path;
                                  
            $AssignmentDetails = ($Assignments | ? {$_.VMName -eq $Name})
            $CollectionName = $AssignmentDetails.Collection;
            $AssignedUser = $AssignmentDetails.User;   
            
            $ComputerName = ($DHCPList | ? {$IPAddress -eq $_.IPAddress}).HostName -replace "\.$Domain$", "";
            
            if($IPAddress -and !$ComputerName) {
                $ComputerName = (Resolve-DnsName $IPAddress -ErrorAction SilentlyContinue).NameHost -replace "\.$Domain$", "";
            }                    
            
            $Uptime = $null;
            if ($_.Uptime -ne 0) {
                $Uptime = "$($_.Uptime.Days.ToString("00")):$($_.Uptime.Hours.ToString("00")):$($_.Uptime.Minutes.ToString("00")):$($_.Uptime.Seconds.ToString("00"))"
            }   
                    
            $Details = New-Object -Type PSCustomObject -Property @{
                
                    'VM Name'         = $Name
                    'Computer Name'   = $ComputerName
                    Server            = $HostServer
                    Memory            = ($_.MemoryStartup / 1GB).ToString("0.00") + " GB"
                    State             = $_.State.ToString()
                    Uptime            = $Uptime
                    'IP Address'      = $IPAddress
                    'Collection Name' = $CollectionName
                    'Assigned User'   = $AssignedUser
                    VHD               = $VHD

            }
            
            $NameMatch = $false;
            $UserMatch = $false;

            foreach ($n in $VMName) {
                if ($Name -like $n -or $ComputerName -like $n) {
                    $NameMatch = $true;
                }
            }

            foreach ($u in $User) {
                if ($AssignedUser -like $u -or $AssignedUser -like "*\$u") {
                    $UserMatch = $true;
                }
            }

            if (!($NameMatch -and $UserMatch)) {
                $Details = $null;
            }
                            
            $script:VMDetails += $Details;
            
            if(!$VHDProperties -or $ExportExcel -or $ExportWord) {   
                $ProgressCount += (30/$VMs.Count);   
            }
            else {
                $ProgressCount += (50/$VMs.Count);
            }
              
        }

    }

    function GetVHDOnServer {
        
        param(
            $CurrentServer,
            $VMs
        )

        $VMsUpdate = Invoke-Command -ComputerName $CurrentServer -ArgumentList $VMs {
                
            $VMs = $args;         

            $VMs |
            % {
                $VMName = $_.'VM Name';
                $DiskNumber = 1;
                $VHDMaster = @();
                $_.VHD | 
                % {
                    $Path = $_;

                    try {                     
                        $VHDList = @();

                        do {
                            $VHDDetails = Get-VHD -Path $Path -ErrorAction Stop;

                            $VHDList += New-Object -Type PSCustomObject -Property @{
                                'VM Name'        = $VMName
                                Server           = $env:COMPUTERNAME
                                Disk             = $DiskNumber
                                Path             = $Path
                                Type             = $VHDDetails.VHDType
                                'Current Size'   = [math]::Round(($VHDDetails.FileSize / 1GB), 2).ToString() + " GB"
                                'Effective Size' = [math]::Round(($VHDDetails.MinimumSize / 1GB), 2).ToString() + " GB"
                                Capacity         = [math]::Round(($VHDDetails.Size / 1GB), 2).ToString() + " GB"

                            }
                            
                            $Parent = $VHDDetails.ParentPath;

                            if ($Parent -ne $null) {
                                $Path = $Parent;
                            }

                        } while ($Parent)
                                           

                    }
                    catch {
                        
                        $Status = 'Corrupted';

                        if ($VHDList) {
                            $Status = 'Corrupt-Snap';
                        }
                                                
                        $vHDList += New-Object -Type PSCustomObject -Property @{
                            'VM Name'        = $VMName
                            Server           = $env:COMPUTERNAME
                            Disk             = $DiskNumber
                            Path             = $Path
                            Type             = $Status
                            'Current Size'   = $null
                            'Effective Size' = $null
                            Capacity         = $null

                        }
                        
                    }                  
                    finally {
                        $DiskNumber++;
                        [array]::reverse($VHDList);
                        $VHDMaster += $VHDList;
                    }
                }

                $_.VHD = $VHDMaster

            }

            return $VMs;

        }

        return $VMsUpdate;

    }

    function VHDProcessing($VMObjects) {
        $VMList = @();
        $ProgressCount = 65;

        $Server | % {
             
            $CurrServer = $_;
            Write-Progress -Activity "Retrieving Hyper-V VHD Details" -Status "Connecting to $CurrServer" -PercentComplete $ProgressCount;
            $ServerVMs = $VMObjects | ? {$_.Server -eq $CurrServer};        
			
            if ($ServerVMs) {
                $VMList += GetVHDOnServer -CurrentServer $_ -VMs $ServerVMs;

            }
            $ProgressCount += (20/$Server.Count);
        }

        $script:VMDetails = $VMList;

    }

    function Main {
           
        if(standardParam -eq 1) {
            return;
        }
        
        $script:VMDetails = @(); 
        GetVMDetails;

        if ($ExportExcel -or $VHDProperties -or $ExportWord) {
            VHDProcessing($VMDetails);
        }

        $ProgressCount = 85;

        Write-Progress -Activity "Retrieving Hyper-V VM Details" -Status "Preparing Results" -PercentComplete $ProgressCount;
        $global:debugVMDetails = $VMDetails;

        if ($VMDetails.Count -eq 0) {
            Write-Output "No matching VMs found.";
            return;
        }

        if (!$ExportExcel -and !$ExportWord) {
            if ($VHDProperties) {

                if($ExtendedProperties) {
                    $VMDetails |
                    % {
                            $_ | Select 'VM Name', 'Computer Name', Server, State, 'IP Address', 'Assigned User', `
                                        'Collection Name', Memory, Uptime | fl;
                            Write-Output "VHD Details";
                            Write-Output "___________"
                            $_.VHD | Select @{n="Disk";e={$_.Disk}}, @{n="Type";e={$_.Type.ToString()}}, `
                                            @{n="Current Size";e={$_.'Current Size'}}, @{n="Effective Size";e={$_.'Effective Size'}}, `
                                            @{n="Capacity";e={$_.Capacity}}, @{n="Path";e={$_.Path}} | ft -AutoSize -Wrap;
                            
                            $ProgressCount += (15/$VMDetails.Count);
                            Write-Progress -Activity "Retrieving Hyper-V VM Details" -Status "Preparing Results" -PercentComplete $ProgressCount;
                        }                       
                }

                else {                       
                    $VMDetails |
                    % {
                            $_ | Select 'VM Name', 'Computer Name', Server, 'IP Address', 'Assigned User' | fl;
                            Write-Output "VHD Details";
                            Write-Output "___________"
                            $_.VHD | Select @{n="Disk";e={$_.Disk}}, @{n="Type";e={$_.Type.ToString()}}, `
                                            @{n="Current Size";e={$_.'Current Size'}}, @{n="Effective Size";e={$_.'Effective Size'}}, `
                                            @{n="Capacity";e={$_.Capacity}}, @{n="Path";e={$_.Path}} | ft -AutoSize -Wrap;
                            
                            $ProgressCount += (15/$VMDetails.Count);
                            Write-Progress -Activity "Retrieving Hyper-V VM Details" -Status "Preparing Results" -PercentComplete $ProgressCount;
                        }                   
                }
                    
            }

            else {
            
                if($ExtendedProperties) {
                    $VMDetails | Select 'VM Name', 'Computer Name', Server, State, 'IP Address', 'Assigned User', `
                                        'Collection Name', Memory, Uptime, VHD | fl;
                }

                else {
                    $VMDetails | Select 'VM Name', 'Computer Name', Server, 'IP Address', 'Assigned User' | ft -AutoSize -Wrap;
                }

            }

        }

        elseif ($ExportExcel) {

            $VMDetails | Select 'VM Name', 'Computer Name', Server, State, 'IP Address', `
                                'Assigned User', 'Collection Name', Memory, Uptime |
                         Export-Excel $ExportExcel -WorkSheetname "VM Info"
            
            $VMDetails.VHD | Select 'VM Name', `
                                    'Server', `
                                    'Disk', `
                                    @{n="Type";e={$_.Type.ToString()}}, `
                                    'Current Size', `
                                    'Effective Size', 
                                    @{n="Capacity";e={$_.Capacity}}, `
                                    @{n="Path";e={$_.Path}} |
                              Export-Excel $ExportExcel -WorkSheetname "VHD Details"

        }

        elseif ($ExportWord) {
           
            $SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
            $Word = New-Object -ComObject Word.Application
            $Word.Visible = $false
            $Document = $Word.Documents.Add()
            $Selection = $Word.Selection

            foreach ($VM in $VMDetails) {

                $Table = $Selection.Tables.add(
                    $Selection.Range, 9, 2,
                    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
                    [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                )
    
                $Table.AllowPageBreaks = 0;

                $Table.Style = "Light Shading - Accent 1"
                $Table.cell(1,1).range.Bold = 1
                $Table.cell(1,1).range.text = 'VM Name'
                $Table.cell(2,1).range.Bold = 1
                $Table.cell(2,1).range.text = 'Computer Name'
                $Table.cell(3,1).range.Bold = 1
                $Table.cell(3,1).range.text = 'Server'
                $Table.cell(4,1).range.Bold = 1
                $Table.cell(4,1).range.text = 'State'
                $Table.cell(5,1).range.Bold = 1
                $Table.cell(5,1).range.text = 'IP Address'
                $Table.cell(6,1).range.Bold = 1
                $Table.cell(6,1).range.text = 'Assigned User'
                $Table.cell(7,1).range.Bold = 1
                $Table.cell(7,1).range.text = 'Collection Name'
                $Table.cell(8,1).range.Bold = 1
                $Table.cell(8,1).range.text = 'Memory'
                $Table.cell(9,1).range.Bold = 1
                $Table.cell(9,1).range.text = 'Uptime'

                $Table.cell(1,2).range.Bold = 0
                $Table.cell(1,2).range.text = $VM.'VM Name'
                $Table.cell(2,2).range.text = $VM.'Computer Name'
                $Table.cell(3,2).range.text = $VM.Server
                $Table.cell(4,2).range.text = $VM.State
                $Table.cell(5,2).range.text = $VM.'IP Address'
                $Table.cell(6,2).range.text = $VM.'Assigned User'
                $Table.cell(7,2).range.text = $VM.'Collection Name'
                $Table.cell(8,2).range.text = $VM.Memory
                $Table.cell(9,2).range.text = $VM.Uptime

                $Word.Selection.Start= $Document.Content.End
                $Selection.TypeParagraph();
    
                $Table = $Selection.Tables.add(
                    $Selection.Range, ($VM.VHD.Count + 1),6,
                    [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord8TableBehavior
                )

                $Table.AllowPageBreaks = 0;

                $Table.Style = "Medium Shading 1 - Accent 1"

                $Table.cell(1,1).range.Bold = 1
                $Table.cell(1,1).range.text = "Disk"
                $Table.cell(1,2).range.Bold = 1
                $Table.cell(1,2).range.text = "Type"
                $Table.cell(1,3).range.Bold = 1
                $Table.cell(1,3).range.text = "Current Size"
                $Table.cell(1,4).range.Bold = 1
                $Table.cell(1,4).range.text = "Effective Size"
                $Table.cell(1,5).range.Bold = 1
                $Table.cell(1,5).range.text = "Capacity"
                $Table.cell(1,6).range.Bold = 1
                $Table.cell(1,6).range.text = "Path"

                for ($i = 0; $i -lt ($VM.VHD.Count); $i++) {
                    $Table.cell(($i+2), 1).range.text = $VM.VHD[$i].Disk.ToString()
                    $Table.cell(($i+2), 2).range.text = $VM.VHD[$i].Type.ToString()
                    $Table.cell(($i+2), 3).range.text = $VM.VHD[$i].'Current Size'
                    $Table.cell(($i+2), 4).range.text = $VM.VHD[$i].'Effective Size'
                    $Table.cell(($i+2), 5).range.text = $VM.VHD[$i].Capacity
                    $Table.cell(($i+2), 6).range.text = $VM.VHD[$i].Path
                }
    
                $Table.Columns[1].SetWidth(30.2, [Microsoft.Office.Interop.Word.WdRulerStyle]::wdAdjustNone);
                $Table.Columns[2].SetWidth(75, [Microsoft.Office.Interop.Word.WdRulerStyle]::wdAdjustNone);
                $Table.Columns[3].SetWidth(60, [Microsoft.Office.Interop.Word.WdRulerStyle]::wdAdjustNone);
                $Table.Columns[4].SetWidth(60, [Microsoft.Office.Interop.Word.WdRulerStyle]::wdAdjustNone);
                $Table.Columns[5].SetWidth(60, [Microsoft.Office.Interop.Word.WdRulerStyle]::wdAdjustNone);
                $Table.Columns[6].SetWidth(180, [Microsoft.Office.Interop.Word.WdRulerStyle]::wdAdjustNone);

                $Word.Selection.Start = $Document.Content.End
                $Selection.TypeParagraph()

                $ProgressCount += (15/$VMDetails.Count);
                Write-Progress -Activity "Retrieving Hyper-V VM Details" -Status "Preparing Results" -PercentComplete $ProgressCount;

            }

            $Document.saveas($ExportWord, $SaveFormat::wdFormatDocumentDefault)
            $Document.close()
            $Word.quit()

        }

    }
    
    Main;

}

# SIG # Begin signature block
# MIIMIQYJKoZIhvcNAQcCoIIMEjCCDA4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUgOcRezE+Ej0mZkZiTaHMPrsg
# 7oOgggoMMIIE7DCCBFWgAwIBAgIKE7/jkQAAAAAADjANBgkqhkiG9w0BAQUFADBS
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
# BgkqhkiG9w0BCQQxFgQUjgvIpHYbcjux6sAYEZ/vyZBi9RAwDQYJKoZIhvcNAQEB
# BQAEgYALH+gHVHeHSQyiVednPCxh/5mZhNVLz99/m4W4PX1KfD+y0I2kmBPi8lZb
# SCOVqT34HvizUuk9xM/vj93cjS2iDivxOCsjF5qEYWD2pXb37lUNavg4Wlng+JBr
# XWwB2UvRspEHHBOsF3R/6OjuvXmxo6peMM3c2dl8LLdU6082HQ==
# SIG # End signature block
