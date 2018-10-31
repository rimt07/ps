Clear-Host

If (-NOT 
	([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
	Write-Warning “No tienes permisos de administrador, por favor corre el script como administrador!”
	#Break
}

Import-Module Webadministration

function Get-AppPoolProperties($Name) {
	#$pool = Get-Item  "IIS:\AppPools\"+ $Name| Select-Object *
    $pool = (Get-Item  ("IIS:\AppPools\"+  $Name)| Select-Object *)
	return $pool
}

function ConvertTo-Json20([object] $item){
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer
    return $ps_js.Serialize($item)
}

function ConvertFrom-Json20([object] $item){ 
    add-type -assembly system.web.extensions
    $ps_js=new-object system.web.script.serialization.javascriptSerializer

    #The comma operator is the array construction operator in PowerShell
    return ,$ps_js.DeserializeObject($item)
}

function Get-ComputerInformation($Computer)
{
	$SysInfo = @()
	$HddInfo =@()

    
    $computerSystem = gwmi Win32_ComputerSystem -Computer $Computer
    $mem = gwmi Win32_PhysicalMemory
    $disk = gwmi Win32_DiskDrive
    $computerBIOS = gwmi Win32_BIOS -Computer $Computer
    $computerOS = gwmi Win32_OperatingSystem -Computer $Computer
    $computerCPU = gwmi Win32_Processor -Computer $Computer
    $computerHDD = gwmi win32_logicaldisk -filter "drivetype = 3" 
    
    $SerialNumber = $computerBIOS | select serialnumber
    $mem = $computerSystem | select NumberOfProcessors,Name,Model
               
    $ServerType = "Physical” # Assume "physical machine" unless resource has "vmware" in the value or a "dash" in the serial #                 

    if ($SerialNumber -like "*-*" -or $SerialNumber -like "*VM*" -or $SerialNumber -like "*vm*") { $ServerType = "Virtual" } 
                   

    $system = New-Object -TypeName PSObject
    $system | Add-Member -Type NoteProperty -Name Server      -Value $computerSystem.Name
    $system | Add-Member -Type NoteProperty -Name ServerType   -Value $ServerType
    $system | Add-Member -Type NoteProperty -Name Manufacturer  -Value $computerSystem.Manufacturer
    $system | Add-Member -Type NoteProperty -Name Model         -Value $computerSystem.Model
    $system | Add-Member -Type NoteProperty -Name CpuName          -Value ($computerCPU.Name)
    $system | Add-Member -Type NoteProperty -Name CpuCores          -Value ($computerCPU.NumberOfCores)
    $system | Add-Member -Type NoteProperty -Name CpuProcessors -Value $mem.NumberOfProcessors
    $system | Add-Member -Type NoteProperty -Name CpuManufacturer -Value ($computerCPU.Manufacturer)
    $system | Add-Member -Type NoteProperty -Name Name          -Value ($computerCPU.Name)
    $system | Add-Member -Type NoteProperty -Name Memory        -Value ("{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB")
    $system | Add-Member -Type NoteProperty -Name SO            -Value ($computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion)
   
    

    foreach($hdd in $computerHDD)
    {
        $itemHdd = New-Object -TypeName PSObject
        $itemHdd | Add-Member -Type NoteProperty -Name Id         -Value $hdd.DeviceID
        $itemHdd | Add-Member -Type NoteProperty -Name FreeSpace  -Value ("{0:N2}" -f ($hdd.FreeSpace/1GB) + "GB")
        $itemHdd | Add-Member -Type NoteProperty -Name Space  -Value ("{0:N2}" -f ($hdd.Size/1GB) + "GB")
        #$itemHdd = $itemHdd | %{“{0}={1}” -f $_.key, $_.value}
        
        $HddInfo += $itemHdd 
    }
    
 
    $HddInfo = $HddInfo | Foreach {"Id=$($_.Id) Space=$($_.Space) FreeSpace=$($_.FreeSpace)"}

    $HddInfo = ($HddInfo | Select-Object -Unique) -Join "||"
   #$HddInfo = ConvertConvertFrom-JsonPSCustomObjectToHash($HddInfo)

    $system | Add-Member -Type NoteProperty -Name Hdd -Value $HddInfo

    $SysInfo += $system
	
	return $SysInfo
}

Function ConvertConvertFrom-JsonPSCustomObjectToHash($obj)
{
    $hash = @{}
     $obj | Get-Member -MemberType Properties | SELECT -exp  "Name" "Value"| % {
                $hash[$_] = ($obj | SELECT -exp $_)
      }
      $hash
}

function Get-WebApps($Computer)
{
   $list = @()
   $listWeb = @()

	foreach ($webapp in Get-ChildItem IIS:\Sites\)
	{
		$name = "IIS:\Sites\" + $webapp.name
		$item = @{}
		$binds = @{}

		$item.WebAppName = $webapp.name

		foreach($Bind in $webapp.Bindings.collection)
		{
			$item.SiteUrl = $Bind.Protocol +'://'+  $Bind.BindingInformation.Split(":")[-1]
			$obj = New-Object PSObject -Property $item
		    $list += $obj
		}

		#Application in site
		$test = Get-Childitem $name | where {$_.Schema.Name -eq 'Application'}
		
	    $test | ForEach-Object { Get-AppPoolProperties($_.applicationPool) }
    
	    $listWeb += $test | Select-Object -Property name, applicationPool,physicalPath,
		@{Name='Server';Expression={$env:COMPUTERNAME}},
        @{Name='RuntimeOfPool';Expression={$pool.managedRuntimeVersion}},
		@{Name='ModeOfPool';Expression={$pool.managedPipelineMode}}#,
		
		
	}

	return $listWeb
}


try{

	$ArrComputers =  "."

	$SystemInfo = @()
	$HddInfo =@()
	$listWebApps = @()
    $outputCollection = @()

    $HostName = get-content env:computername
    $csvPath = "{0}-WebApplicationInventory.csv" -f $HostName

	foreach ($source in $ArrComputers)
	{
		#Set-ExecutionPolicy unrestricted
		$SystemInfo += Get-ComputerInformation($source);

        #$SystemInfo | Format-Table -Auto

        $listWebApps += Get-WebApps($source);

        #$listWebApps | Format-Table -Auto
	
        $SystemInfo | Foreach-Object {
            #Associate objects
            $systemObject = $_
            $FilteredWebObject = $listWebApps | Where-Object {$_.Server -eq $systemObject.Server}
                
            $FilteredWebObject | Foreach-Object {
            
                $outputObject = New-Object -TypeName PSObject
                
                $ConnStrList= @()
                $WcfClientList= @()
                
               try
               {
               
                 if(Test-Path -Path $_.PhysicalPath)
                 {
  
                   $appConfig = [xml](cat ($_.PhysicalPath + "\web.config"))
                
                  
                   $appConfig.configuration.connectionStrings.add | 
                   foreach { $ConnStrList +=  $_.connectionString }
                   
                
                    $nodoEndpointSvcAddress = $($appConfig.configuration.'system.serviceModel'.client.endpoint) |
                    Foreach-Object { 
                    [string]$endpointSvcAddress = $_.address 
                    $WcfClientList +=  $endpointSvcAddress
                     }

                   
                    $ConnStrList = ($ConnStrList | Foreach {"$_"} | Select-Object -Unique) -Join "||"
                    $WcfClientList = ($WcfClientList | Foreach {"$_"} | Select-Object -Unique) -Join "||"
                  }
                  else
                  { 
                     $ConnStrList =($ConnStrList  | Select-Object -Unique)-Join ","
                     $WcfClientList =($WcfClientList  | Select-Object -Unique)-Join ","
                  }
                 }
                 catch
                 {
                     $ConnStrList =($ConnStrList  | Select-Object -Unique)-Join ","
                     $WcfClientList =($WcfClientList  | Select-Object -Unique)-Join ","
                     #Write-Host "config don't found"
                 }
                       
                $outputObject | Add-Member -Type NoteProperty -Name ConnectionString  -Value $ConnStrList 
                $outputObject | Add-Member -Type NoteProperty -Name Endpoints  -Value $WcfClientList 
               
               
                $outputObject | Add-Member -Type NoteProperty -Name Hdd           -Value $systemObject.Hdd
                $outputObject | Add-Member -Type NoteProperty -Name Server        -Value $systemObject.Server   
                $outputObject | Add-Member -Type NoteProperty -Name ServerType    -Value $systemObject.ServerType  
                $outputObject | Add-Member -Type NoteProperty -Name Manufacturer  -Value $systemObject.Manufacturer
                $outputObject | Add-Member -Type NoteProperty -Name Model         -Value $systemObject.Model 
                $outputObject | Add-Member -Type NoteProperty -Name CpuName       -Value $systemObject.CpuName  
                $outputObject | Add-Member -Type NoteProperty -Name CpuManufacturer -Value $systemObject.CpuManufacturer        
                $outputObject | Add-Member -Type NoteProperty -Name CpuCores      -Value $systemObject.CpuCores
                $outputObject | Add-Member -Type NoteProperty -Name CpuProcessors -Value $systemObject.CpuProcessors                            
                $outputObject | Add-Member -Type NoteProperty -Name Memory        -Value $systemObject.Memory
                
                $outputObject | Add-Member -Type NoteProperty -Name SO            -Value $systemObject.SO  
                $outputObject | Add-Member -Type NoteProperty -Name WepAppName    -Value $_.Name
                $outputObject | Add-Member -Type NoteProperty -Name AppPool       -Value $_.applicationPool  
                $outputObject | Add-Member -Type NoteProperty -Name PhysicalPath  -Value $_.PhysicalPath                                          
                $outputObject | Add-Member -Type NoteProperty -Name RuntimeOfPool -Value $_.RuntimeOfPool
                $outputObject | Add-Member -Type NoteProperty -Name ModeOfPool    -Value $_.ModeOfPool 
                
                                                                                                                                   
             #Add the object to the collection
            $outputCollection += $outputObject 
            }
        }
  
    }
		
        $outputCollection   |
        #Format-Table -Auto 
       Export-Csv -path c:\$csvPath -NoTypeInformation
        
}
catch
{
	$ExceptionMessage = 
	"Error in Line: " + 
	$_.Exception.Line + ". " + $_.Exception.GetType().FullName + ": " + $_.Exception.Message + 
	" Stacktrace: "    + $_.Exception.StackTrace

	$ExceptionMessage
}


