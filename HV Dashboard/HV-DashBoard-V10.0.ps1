Function Get-VMStats {

<#
    .Synopsis
        Gets stats on a VM

#>    

    [CmdletBinding()]
    Param (
        [Parameter (Mandatory = $True, ValueFromPipeLine = $True)]
        [Microsoft.HyperV.PowerShell.VirtualMachine]$VM,

        [PSObject]$VMCPUPerf
    )

    Process {
        Write-Verbose "Getting VM infor for $($VM.Name)"

         # ------ Performance counter
        $CPUPerf = $VMCPUPerf | where { ($_.ComputerName).tolower() -eq $VM.Name.tolower() } | Select-Object -Last 72  #| Select-object PercentRunTime | Measure-object -Average -Maximum
        $CPUPerf = $CPUPerf.PercentRunTime | Measure-object -Average -Maximum
			          

		$VMInfo = New-Object -TypeName PSCustomObject -Property (@{
            'Name' = $VM.Name
            'ProcessorCount' = $VM.ProcessorCount
            'MemoryDemand' = $VM.MemoryDemand
            'VMID' = $VM.VMID
            'MinimumRAM' = $VM.MemoryMinimm/1GB
            'MaximumRAM' = $VM.MemoryMaximum/1GB
            'DynamicMemoryEnabled' = $VM.DynamicMemoryEnabled
            'StartUpRAM' = $VM.MemoryStartup/1GB
            'IntegrationServicesVersion' = $VM.IntegrationServicesVersion
            'Uptime' = $VM.UpTime
            'ParentSnapShotID' = $VM.ParentSnapshotId
            'IntegrationServicesState' = $VM.IntegrationServicesState
            '$ResourceMeteringEnabled' = $VM.ResourceMeteringEnabled
            'AvgMemUsage' = 'NA'
            'DiskDetails' = Get-VMHardDiskDrive -VM $VM | Get-VHD | Foreach {
                $VMDiskINfo = $_
         
                $VMDiskInfo | Add-Member -MemberType NoteProperty -Name VolInfo -Value (($_.path).split('\')[2])
	            $VMDiskInfo | Add-Member -MemberType NoteProperty -Name 'VolName' -Value { IF ($_.Path -match "C:\\ClusterStorage") {($_.Path).Split("\")[2]} Else {($_.Path).Substring(0,2)}}
	            $VMDiskInfo | Add-Member -MemberType NoteProperty -Name 'vDiskType' -Value ( ($_.path).split('.')[1] )
		
                $VMDiskInfo
            }
            'SnapShotDate' = 'NA'
            'VNICType' = (Get-VMNetworkAdapter -VMName $VM.Name).IsLegacy
            'IsClustered' = $VM.IsClustered
			'VMCPUPerf' = $CPUPerf
            
        })	

        # ----- I couldn't figure out how to do an if statement in the object creation.
        if ( $VM.ResourceMeteringEnabled ) { 
            $VMInfo.AvgMemUsage = ((Measure-VM -VMName $VM.name).AvgRam)/1024 
        }

        If ($VMDetails.ParentSnapshotId) {
			$VMInfo.SnapSHotDate = ((Get-VMSnapshot -VMName $VM.name).CreationTime | Sort-Object | Select-Object -First 1).ToShortDateString()
		}
		
        
        Write-Output $VMInfo		        
    }
}

#-------------------------------------------------------------------------------------

Function Get-ServerStats {

<#
    .Synopsis
        Builds the bulk of the Dashboard
        
    .Description
        This function is split into two parts.  One runs if the host belongs to a cluster.  The other if it is a standalone host.   

    .Parameter Name
        Name of they Hyper-V HOst

    .parameter Type
        Either Cluster or StandAlone.
#>
    [CmdletBinding()]
    Param(
        [Parameter ( Mandatory = $True,ValueFromPipeline = $True)]
        [String[]]$Name, 
        
      #  [Parameter ( Mandatory = $True)]
        [ValidateSet ( 'Cluster','Standalone' )]
        [String]$Type = 'Standalone',

        [Parameter ( Mandatory = $true ) ]
        [PSObject]$HostCPUPerf,

        [Parameter ( Mandatory = $true ) ]
        [PSObject]$VMCPUPerf
    )
	
    Process {
        Foreach ( $N in $Name ) {
            Write-Verbose "Processing Host $N"
           
            # ----- Local Storage
            $LocalStorage = Get-WmiObject Win32_LogicalDisk -filter "DriveType=3" -computer $Name | Select DeviceID, Size, FreeSpace, @{n="FreeSpacePC";e={[int]($_.FreeSpace/$_.Size*100)}}, @{n="VHDXAllocatedSpace";e={0}}, @{n="VHDXActualUsage";e={0}}, @{n="VolumeHealthCode";e={[int]"0"}}
		
            # ----- Memory
            $hostDetails = Get-WmiObject -Class win32_OperatingSystem -ComputerName $Name |Select @{Label="TotalPhysicalMemory"; Expression={[int]($_.TotalVisibleMemorySize/1048576)}}, @{Label="AvailableMemory"; Expression={[int]($_.FreePhysicalMemory/1048576)}}, @{Label="AvailablePhysicalMemoryPC"; Expression={[int](($_.FreePhysicalMemory/$_.TotalVisibleMemorySize)*100)}}	     

            # ----- Processor
            $procDetails = Get-WmiObject -Class Win32_Processor -Computername $Name

            # ------ Performance counter
           $CPUPerf = $HostCPUPerf | where { ($_.ComputerName).tolower() -eq $N.tolower() } | Select-Object -Last 72  #| Select-object PercentRunTime | Measure-object -Average -Maximum
            $CPUPerf = $CPUPerf.PercentRunTime | Measure-object -Average -Maximum
			
            # ----- CPU Performance
            #$HostCPUPerf = import-csv $HostCPUFile

            $HostServer = New-Object -TypeName PSCustomObject -Property (@{
                'ComputerName' = $N
                'TotalPhysicalMemory' = $hostDetails.TotalPhysicalMemory
                'AvailableMemory' = $HostDetails.AvailableMemory
                'AvailablePhysicalMemoryPercent' = $hostDetails.AvailablePhysicalMemoryPC
                'LocalStorage' = $LocalStorage
                'Processors' = ($procDetails.DeviceID).Count
                'ProcessorCore' = ($procDetails.numberofcores |Measure-Object -Sum).sum
                'LogicalProcessors' = ($procDetails.numberoflogicalprocessors |Measure-Object -Sum).sum
                'VMDetails' = ( Get-VM | Get-VMStats -VMCPUPerf $VMCPUPerf )
                'HostCPUPerf' = $CPUPerf 
            })
      
            Write-Output $HostServer
        }
   }
}

#-------------------------------------------------------------------------------------

Function fWriteHtmlHeader { 

	param($FileName) 

	$date = ( get-date ).ToString('yyyy/MM/dd') 
	Add-Content $FileName "<html>" 
	Add-Content $FileName "<head>" 
	Add-Content $FileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
	Add-Content $FileName '<title>Hyper-V Dashboard</title>' 
	Add-Content $FileName '<STYLE TYPE="text/css">' 
	Add-Content $FileName  "<!--" 
	Add-Content $FileName  "td {" 
	Add-Content $FileName  "font-family: Tahoma;" 
	Add-Content $FileName  "font-size: 11px;" 
	Add-Content $FileName  "border-top: 2px solid #999999;" 
	Add-Content $FileName  "border-right: 2px solid #999999;" 
	Add-Content $FileName  "border-bottom: 2px solid #999999;" 
	Add-Content $FileName  "border-left: 2px solid #999999;" 
	Add-Content $FileName  "}" 
	Add-Content $FileName  "body {" 
    Add-Content $FileName  "margin-left: 5px;" 
	Add-Content $FileName  "margin-top: 5px;" 
	Add-Content $FileName  "margin-right: 5px;" 
	Add-Content $FileName  "margin-bottom: 5px;" 
	Add-Content $FileName  "" 
	Add-Content $FileName  "table {" 
	Add-Content $FileName  "border: thin solid #000000;" 
	Add-Content $FileName  "}" 
	Add-Content $FileName  "-->" 
	Add-Content $FileName  "</style>" 
	Add-Content $FileName "</head>" 
	Add-Content $FileName "<body>" 
	Add-Content $FileName  "<table width='100%'>" 
	Add-Content $FileName  "<tr bgcolor='#2F0B3A'>" 
	Add-Content $FileName  "<td colspan='30' height='20' align='center'>" 
	Add-Content $FileName  "<font face='tahoma' color='#FFFF00' size='4'><strong>Hyper-V - VM Dashboard -  $date</strong></font>" 
	Add-Content $FileName  "</td>" 
	Add-Content $FileName  "</tr>" 
	Add-Content $FileName  "</table>" 
}

#-------------------------------------------------------------------------------------

Function fWriteSubHeadingClusterOrStandAlone {

	Param ($FileName, $cname)

   	Add-Content $FileName  "<table width='100%'>" 
	Add-Content $FileName  "<tr colspan='1' height='20' align='center' bgcolor='#000000'>" 
	Add-Content $FileName  "<td width = '100%' color='#000000' size='2' align=center><font color='#FFFC00'><strong>$cname</strong></font></td>" 
	Add-Content $FileName  "</tr>" 
	Add-Content $FileName  "</table>" 
}

#-------------------------------------------------------------------------------------

Function Write-ServerNode {

    [CmdletBinding()]
    Param (
        [Parameter (Mandatory = $True)]
        [String]$FileName,

        [Parameter (Mandatory = $True, ValueFromPipeline = $True )]
        [PSObject]$HostServer
    )

	#Param ($FileName, $nodeName, $TotMem, $AvailMem, $AvailMemPC, $hostMemHealth,[PSObject]$HostCPUPerf)

    Begin {
         

        [Array]$WarningLevel = "#77FF5B","#FFF632","#FF6B6B","#FF0040"
    }

    Process {
        Write-verbose "Adding $($HostServer.ComputerName) to ResultFile"

        
        If ( $HostServer.CPUPerf.AVG -ge 90 ) { $CPUAVGHealth =  2 }
            Elseif ( $CPUPerf.AVG -ge 80 -and $CPUPerf.AVG -lt 90 ) { $CPUAVGHealth = 1 }
                Else { $CPUAVGHealth = 0 }

        If ( $HostServer.CPUPerf.MAX -ge 90 ) { $CPUMAXHealth =  2 }
            Elseif ( $CPUPerf.MAX -ge 80 -and $CPUPerf.MAX -lt 90 ) { $CPUMAXHealth = 1 }
                Else { $CPUMAXHealth = 0 }

        If (($HostServer.AvailablePhysicalMemoryPercent -le "10") -OR ($HostServer.AvailableMemory -lt "10")) {
			$hostMemHealth = "3"
		}
		ElseIf ((($HostServer.AvailablePhysicalMemoryPercent -le "20") -And ($HostServer.AvailablePhysicalMemoryPercent -gt "10")) -OR (($HostServer.AvailableMemory -lt "20") -And ($HostServer.AvailableMemory -gt "5"))) {
			$hostMemHealth = 2
		}
		ElseIf ((($HostServer.AvailablePhysicalMemoryPercent -le "30") -And ($HostServer.AvailablePhysicalMemoryPercent -gt "20")) -OR (($HostServer.AvailableMemory -lt "30") -And ($HostServer.AvailableMemory -gt "10"))) {
			$hostMemHealth = 1
		}
		Else {
		    $hostMemHealth = 0
		}

        Add-Content $FileName  "<table width='100%'>" 
	    Add-Content $FileName  "<tr height='20' bgcolor='#000000'>" 
	    Add-Content $FileName  "<td width = '40%' size='3' align=center><font color='White'><strong>$($HostServer.ComputerName)</strong></Font></td>" 
        Add-Content $FileName "<td width = '8%' align=center><font color='White'>Max CPU</Font></td>"
        Add-Content $FileName "<td width = '4%' align=center><Strong><font Size='4' color='$($WarningLevel[$CPUMAXHealth])'>$("{0:N0}" -f$HostServer.CPUPerf.Maximum) %</Font></strong></td>"
        Add-Content $FileName "<td width = '8%' align=center><font color='White'>Avg CPU</Font></td>"
        Add-Content $FileName "<td width = '4%' align=center><Strong><font Size='4' color='$($WarningLevel[$CPUAVGHealth])'>$("{0:N0}" -f$HostServer.CPUPerf.Average) %</Font></strong></td>"
	    Add-Content $FileName "<td width='8%' align=center><font color='White'>Total Memory</Font></td>"
	    Add-Content $FileName "<td width='4%' align=center><font color='White'>$($HostServer.TotalPhysicalMemory) GB</Font></td>"
	    Add-Content $FileName "<td width='8%' align=center><font color='White'>Available Memory</Font></td>"
	    Add-Content $FileName "<td width='4%' align=center><font color='White'>$($HostServer.AvailableMemory) GB</Font></td>"
	    Add-Content $FileName "<td width='8%' align=center><font color='White'>Available Memory (%)</Font></td>"
	    Add-Content $FileName "<td width='4%' align=center><strong><font size='4' color='$hostMemHealth'>$HostServer.AvailablePhysicalMemoryPercent % </font></strong></td>"
	    Add-Content $FileName  "</tr>"
	    Add-Content $FileName  "</table>"
    
        

        $HostServer.VMDetails | Write-VMInfo -FileName $FileName -HostMemHealth $HostMemHealth -Verbose

    }
}

#-------------------------------------------------------------------------------------

Function Write-VMInfo {

<#
    .Synopsis
        Outputs the stats info about each VM
#>
 

    [CmdletBinding()]
    param (
        [Parameter (Mandatory = $True)]
        [String]$FileName,

        [Parameter (Mandatory = $True, ValueFromPipeLine = $True ) ]
        [PSObject]$VM,

        [Int]$HostMemHealth
    )	

#Param($FileName, $vmname, $utime, $ic, $clusterrole, $vProc, $Startmem, $MinMem, $MaxMem, $AvgMem, $hostmemhealth, $vd1storage, $vd1, $vdu1, $vd1FP, $vd1StorageHealth, $vdtype1, $vd2Storage, $vd2, $vdu2, $vd2fp, $vd2StorageHealth, $vdtype2, $vd3Storage, $vd3, $vdu3, $vd3fp, $vd3StorageHealth, $vdtype3, $vNetworkInterfaceType, $SSDate, $ICStatus, $vhealth, 
    #    [PSObject]$VMCPUPerf
    #)

    Begin {
        [Array]$WarningLevel = "#77FF5B","#FFF632","#FF6B6B","#FF0040"
    
        # ----- Create Headers
        Add-Content $FileName  "<table width='100%'>"
 	    Add-Content $FileName "<tr bgcolor=#BE81F7>" 
	    Add-Content $FileName "<td width='8%' align=center>VM</td>"
	    Add-Content $FileName "<td width='5%' align=center>Up-Time</td>"
	    Add-Content $FileName "<td width='5%' align=center>IC Version</td>"
	    Add-Content $FileName "<td width='4%' align=center>Clustered</td>"
	    Add-Content $FileName "<td width='2%' align=center>vProcessor</td>"
        Add-Content $FileName "<td width='4%' align=center>vProc % Max</td>"
        Add-Content $FileName "<td width='4%' align=center>vProc % Avg</td>"
	    Add-Content $FileName "<td width='4%' align=center>vRAM-StartUp</td>"
    	Add-Content $FileName "<td width='4%' align=center>vRAM-Min</td>"
    	Add-Content $FileName "<td width='4%' align=center>vRAM-Max</td>"
    	#Add-Content $FileName "<td width='6%' align=center>vRAM-Avg</td>"
	    Add-Content $FileName "<td width='6%' align=center>vDisk1-Storage</td>"
	    Add-Content $FileName "<td width='6%' align=center>vDisk1-Allocated</td>"
	    Add-Content $FileName "<td width='6%' align=center>vDisk1-Usage</td>"
	    #Add-Content $FileName "<td width='6%' align=center>vDisk1-FP</td>"
	    Add-Content $FileName "<td width='6%' align=center>vDisk2-Storage</td>"
	    Add-Content $FileName "<td width='6%' align=center>vDisk2-Allocated</td>"
	    Add-Content $FileName "<td width='6%' align=center>vDisk2-Usage</td>"
	    #Add-Content $FileName "<td width='6%' align=center>vDisk2-FP</td>"
	    #Add-Content $FileName "<td width='6%' align=center>vNic</td>"
	    #Add-Content $FileName "<td width='6%' align=center>FirstSnapShotDate</td>"
        Add-Content $FileName "<td width='6%' align=center>vDisk3-Storage</td>"
	    Add-Content $FileName "<td width='6%' align=center>vDisk3-Allocated</td>"
	    Add-Content $FileName "<td width='6%' align=center>vDisk3-Usage</td>"
	    Add-Content $FileName "</tr>"
    }

    Process {
        
        If ( $VM.CPUPerf.Average -ge 90 ) { $CPUAVGHealth =  2 }
            Elseif ( $VM.CPUPerf.Average -ge 80 -and $VM.CPUPerf.Average -lt 90 ) { $CPUAVGHealth = 1 }
                Else { $CPUAVGHealth = 0 }

        If ( $VM.CPUPerf.Maximum -ge 90 ) { $CPUMAXHealth =  2 }
            Elseif ( $VM.CPUPerf.Maximum -ge 80 -and $VM.CPUPerf.Maximum -lt 90 ) { $CPUMAXHealth = 1 }
                Else { $CPUMAXHealth = 0 }



        Add-Content $FileName "<tr bgcolor=#77FF5B>"
	    Add-Content $FileName "<td width='8%' align=center>$($VM.Name)</td>" 
	    Add-Content $FileName "<td width='5%' align=center>$("{0:N0}" -f $VM.Uptime)</td>"
	    If ($VM.IntegrationServicesState -match "Update required") {
		    Add-Content $FileName "<td width='5%' align=center BGCOLOR='#0044FF'><font color='White'><strong>$($VM.IntegrationServicesVersion)</strong></font></td>"
	    }
	    Else {
		    Add-Content $FileName "<td width='5%' align=center>$($VM.IntegrationServicesVersion)</td>"
	    }
	    Add-Content $FileName "<td width='2%' align=center>$($VM.IsClustered)</td>" 
	    Add-Content $FileName "<td width='2%' align=center>$($VM.ProcessorCount)</td>"
        #Add-Content $FileName "<td width='4%' align=center><Strong><font Size='4' color='$($WarningLevel[$CPUMAXHealth])'>$("{0:N0}" -f$CPUPerf.Maximum) %</td>"
        #Add-Content $FileName "<td width='4%' align=center><Strong><font Size='4' color='$($WarningLevel[$CPUMAXHealth])'>$("{0:N0}" -f$CPUPerf.Average) %</td>"
    
        # ----- if Max is null (VM is off) set to Null
        if ( $VM.MaximumRAM -eq $Null ) { $VM.CPUPerf.Max = 0 } 
        Add-Content $FileName "<td width='4%' BGColor=$($WarningLevel[$CPUMAXHealth]) align=center>$("{0:N0}" -f $VM.CPUPerf.Maximum) %</td>"

        # ----- if Max is null (VM is off) set to Null
        if ( $VM.AvgMemUsage -eq $Null ) { $VM.AvgMemusage = 0 }
        Add-Content $FileName "<td width='4%' BGColor=$WarningLevel[$CPUAVGHealth] align=center>$("{0:N0}" -f $VM.CPUPerf.Average) %</td>"

	    Add-Content $FileName "<td width='5%' BGCOLOR=$hostmemhealth align=center>$($VM.StartUPRAM) GB</td>"
    #	If ($VM.MinimumRAM -eq "DM Disabled") {
		    Add-Content $FileName "<td width='6%' align=center>$("{0:N0}" -f ($VM.MinimumRAM/1GB)) GB</td>"
            Write-Verbose $($VM.MaximumRAM/1GB)
    	    Add-Content $FileName "<td width='6%' align=center>$("{0:N0}" -f ($VM.MaximumRAM/1GB)) GB</td>"
    #	}
    #	Else {
    #    	Add-Content $FileName "<td width='6%' align=center>$MinMem GB</td>"
    #    	Add-Content $FileName "<td width='6%' align=center>$MaxMem GB</td>"
    #	}


        # ----- Loop thru VHD info
        Foreach ( $VHD in $VM.DiskDetails ) {
            # ----- Note: VHD health is actually the health of the storage it reside upon.
            Add-Content $FileName "<td width='6%' BGCOLOR='$vd1StorageHealth' align=center>$vd1Storage</td>"
	        If (($VHD.vDiskType -like "vhd*") -OR ($VHD.vDiskType -like "avhd*")) {
		        Add-Content $FileName "<td width='6%' BGCOLOR='#0044FF' align=center><font color='White'><strong>$("{0:N0}" -f ($VHD.FileSize/1GB)) GB</strong></font></td>"
		        Add-Content $FileName "<td width='6%' BGCOLOR='#0044FF' align=center><font color='White'><strong>$("{0:N0}" -f ($VHD.FragmentationPercentage/1GB)) GB</strong></font></td>"
		    }
	        Else {
		        Add-Content $FileName "<td width='6%'  align=center>$("{0:N0}" -f ($VHD.FileSize/1GB)) GB</td>"
		        Add-Content $FileName "<td width='6%'  align=center>$("{0:N0}" -f ($VHD.FragmentationPercentage/1GB)) GB</td>"
		    }
        }

	    


	 
	    #If ($vNetworkInterfaceType -eq "False")
		    #{
		    #Add-Content $FileName "<td width='6%' BGCOLOR='#0044FF' align=center><font color='White'><strong>Legacy</strong></font></td>"
		    #}
	    #Else
		    #{
		    #Add-Content $FileName "<td width='6%' align=center>Synthetic</td>"
		    #}
	    #Add-Content $FileName "<td width='6%' align=center>$SSDate</td>"
	    Add-Content $FileName "</tr>"

    }
}

#-------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------

$ResultFile = 'c:\temp\hvreport.html'
$HostCPUPerf = import-CSV c:\temp\CloudHostCPU.csv
$VMCPUPerf = import-CSV c:\temp\CloudVMCPU.csv

$Servers = Get-Content c:\temp\Servers.txt

$HostServers = $Servers | Get-ServerStats -HostCPUPerf $HostCPUPerf -VMCPUPerf $VMCPUPerf -Verbose

# ----- Build the Report
fWriteHtmlHeader $ResultFile
fWriteSubHeadingClusterOrStandAlone $ResultFile ("Standalone Node - " +$name)

$HostServers | Write-ServerNode -FileName $ResultFile -verbose



#If ($vDiskcount -gt "1") {
#	fWriteVMInfo  $ResultFile $vmDetails.name $vmDetails.UpTime $vmDetails.IntegrationServicesVersion "NA" $vmDetails.ProcessorCount $vmDetails.StartupRam $vmDetails.MinimumRam $vmDetails.MaximumRam $AvgMemUsage $HostMemoryHealth $DiskDetails[0].VolName ([int]($DiskDetails.AllocatedSize[0]/1073741824)) ([int]($DiskDetails.CurrentUsage[0]/1073741824)) $DiskDetails[0].FragmentationPercentage $WarningLevel[$DiskDetails[0].vDiskVolumeHealth] $DiskDetails[0].vDiskType $DiskDetails[1].VolName ([int]($DiskDetails.AllocatedSize[1]/1073741824)) ([int]($DiskDetails.CurrentUsage[1]/1073741824)) $DiskDetails[1].FragmentationPercentage $WarningLevel[$DiskDetails[1].vDiskVolumeHealth] $DiskDetails[1].vDiskType $vNicType $SnapShotDate $vmDetails.IntegrationServicesState $WarningLevel[$vmHealth] $VMCPUPerf
#}
#Else {                   
#	fWriteVMInfo  $ResultFile $vmDetails.name $vmDetails.UpTime $vmDetails.IntegrationServicesVersion "NA" $vmDetails.ProcessorCount $vmDetails.StartupRam $vmDetails.MinimumRam $vmDetails.MaximumRam $AvgMemUsage $WarningLevel[$HostMemoryHealth] $DiskDetails[0].VolName ([int]($DiskDetails.AllocatedSize[0]/1073741824)) ([int]($DiskDetails.CurrentUsage[0]/1073741824)) $DiskDetails[0].FragmentationPercentage $WarningLevel[$DiskDetails[0].vDiskVolumeHealth] $DiskDetails[0].vDiskType "NA" "NA" "NA" "NA" "NA" "NA" $vNicType $SnapShotDate $vmDetails.IntegrationServicesState $WarningLevel[0] -VMCPUPerf $VMCPUPerf        
#}