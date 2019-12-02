<#
.SYNOPSIS
                MS Exchange- Microsoft Exchange Diagnostics service crashing post CU upgrade.

.DESCRIPTION
                The script will be run and gather real time data which will helpful for the technician at the time of troubleshooting:
                    1. Checking Windows Scheduler Service status
                    2. Checking if required 4 templates are present on the machine or not
                    3. Checking if .blg extention files are present in the required location or not
                    4. checking if ExchangeDiagnosticsDailyPerformanceLog and ExchangeDiagnosticsPerformanceLog tasks are present or not
                

.NOTES    
                Name: MicrosoftExchangeDiagnosticsserviceCrashingPostCUUpgrade.ps1
				Author: <Vinit Keny - GRT, Continuum>
				Version: 1.0
				Created On: 11-18-2019 
  
.PARAMETERS 
                NA.

#>

<# Architecture check started and PS changed to the OS compatible #>
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
	Write-Warning "Excecuting the script under 64 bit powershell"
	if ($myInvocation.Line) {
		& "$env:systemroot\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
	} else {
		& "$env:systemroot\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile -File "$($myInvocation.InvocationName)" $args
	}
	exit $lastexitcode
}
<#Architecture check completed #>

<# Compatibility check if found incompatible will exit #>
try {
	[double]$OSVersion = [Environment]::OSVersion.Version.ToString(2)
	$PSVersion = (Get-Host).Version
	if (($OSVersion -lt 6.1) -or ($PSVersion.Major -lt 2))
	{
		Write-Output "[MSG: System is not compatible with the requirement. Either machine is below Windows 7 / Windows 2008R2 or Powershell version is lower than 2.0]"
		exit
	}
} catch { Write-Output "[MSG: ERROR : $_.Exception.message]" }
<# Compatibility Check Code Ends #>

#>



<#

#Function to check if the server is an exchange server or not
#Checks whether MSExchangeServiceHost is present or not
#Looks for the PSSnapin of Exchange, if not present installs the same

#>


function check_exchange {

    $ErrorActionPreference = "SilentlyContinue"

    if (!(Get-Service -name MSExchangeADTopology -ComputerName $env:COMPUTERNAME)){

		Return $true
	}
	
	if (!(Get-PSSnapin *Exchange*)){
		Add-PSSnapin *Exchange*
	}
}

<# The below function will check if the Windows Scheduler Service is present on the server and is set to Automatic/Manual and started #>

function check_service{
	try{
	    $srv = Get-Service -Name schedule, pla -ComputerName $env:COMPUTERNAME -ErrorAction Stop| select DisplayName, Name, StartType, Status | Format-List | Out-String
    }
	catch{
        $srv = "The schedule/pla service is not existing on this server."
        $srv += "[MSG: ERROR : $($_.Exception.message)]"
    }

    return $srv
	
}


<# The below function will check if the templates key are present under PLA registry path #>
function check_templates{ 
	try{
    $prop = Get-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\PLA\Templates" -ErrorAction Stop 
	if($prop.'{7478EF61-8C46-11d1-8D99-00A0C913CAD4}' -and $prop.'{7478EF62-8C46-11d1-8D99-00A0C913CAD4}' -and $prop.'{7478EF62-8C46-11d1-8D99-00A0C913CAD4}'){
	
$test = @"
"Below template files are present in the registry path 'HKLM\SOFTWARE\Microsoft\PLA\Templates'"
									#Key = #Value                   
"{7478EF61-8C46-11d1-8D99-00A0C913CAD4}" = $($prop.'{7478EF61-8C46-11d1-8D99-00A0C913CAD4}')
"{7478EF62-8C46-11d1-8D99-00A0C913CAD4}" = $($prop.'{7478EF62-8C46-11d1-8D99-00A0C913CAD4}')
"{7478EF63-8C46-11d1-8D99-00A0C913CAD4}" = $($prop.'{7478EF63-8C46-11d1-8D99-00A0C913CAD4}')
"@
	}
    else{
        $test = "The template files are not present in the registry path 'HKLM:SOFTWARE\Microsoft\PLA\Templates' which are essential to fix the issue"
    }
	
    return $test
}

catch{
    return "[MSG: ERROR : $($_.Exception.message)]"
}

}

<# Below function will check if .blg extention files are present in the required location or not #>


function blg{
   try{
		$blgpath = "$env:ExchangeInstallPath\Logging\Diagnostics\DailyPerformanceLogs\*.blg"
		return (Test-Path -Path $blgpath)		
       }
   catch{
        return $false
   }
}
		
	function blg_test {
		try{
		if (blg) {
			$out = "The files with the extension '.blg' are present"
			Get-ChildItem -path "$env:ExchangeInstallPath\Logging\Diagnostics\DailyPerformanceLogs\*.blg"  |`
				ForEach-Object { $out += @"
		`n
LastWriteTime : $($_.LastWriteTime)
FileName      : $($_.Name)
"@
			}
		}
		else {
			$out = "The files with the extension '.blg' are not present"
		}
		
		return $out
   }
        
   catch{
        $out = "Some error occurred"
        $out += "[MSG: ERROR : $($_.Exception.message)]"
        }	
}

<# Below function will check if ExchangeDiagnosticsDailyPerformanceLog and ExchangeDiagnosticsPerformanceLog tasks are present or not #>
function task_return {
    
    if ($OSversion -ge 6.2)
    {
    $task = Get-ScheduledTask -TaskName ExchangeDiagnosticsDailyPerformanceLog,ExchangeDiagnosticsPerformanceLog -ErrorAction SilentlyContinue | Select TaskName, State
    if($task){
        $task | Format-List | Out-String
    }
    else{
        Write-Output "ExchangeDiagnosticsDailyPerformanceLog/ExchangeDiagnosticsPerformanceLog is not present on the system"
        }
    }else {
 
        $task = schtasks.exe /query /fo csv | ConvertFrom-Csv |  Where-Object{$_.TaskName -like "*ExchangeDiagnosticsDailyPerformanceLog" -or $_.TaskName -like "*ExchangeDiagnosticsPerformanceLog"}|  Select TaskName, Status
    
        if($task){
            $task | Format-List | Out-String
        }
        else{
            Write-Output "ExchangeDiagnosticsDailyPerformanceLog/ExchangeDiagnosticsPerformanceLog is not present on the system"
        }
    
     }

}


try{

    $check_exchange = check_exchange
	if($check_exchange){
		Write-Output "The Exchange server application is not installed on this machine"
		exit
	}

    Write-Output "==================================================================`n"
    task_return
    Write-Output "==================================================================`n"
    check_templates
    Write-Output "==================================================================`n"
    blg_test
    Write-Output "==================================================================`n"
    check_service    

}

catch{
    Write-Output "Some error occurred"
}
