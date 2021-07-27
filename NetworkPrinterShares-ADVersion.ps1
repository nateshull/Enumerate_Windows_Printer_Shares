#*****************************************************
#Powershell v5.1
#Copyright Nate Shull 2021-07-26
#Apache License Version 2.0
#Gets a list of computers from AD 
#checks if the computer is reachable
#then looks for any shared printers on the computer
#exports to csv
#*****************************************************

#filter to specific ous
$oufilter = "OU=orgunit, DC=domain, DC=com" 

#get script path and put csv file in same path. You can change the output with the csvfilepath variable. 
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$csvfilepath = "$ScriptDir\PrintServerlist.csv"


#get list of computers from active directory
#could be swapped for a csv path if you want columns: Name, Description
$computerlist = Get-ADComputer -Filter * -SearchBase $oufilter -Properties Name,Description

#results compiled to this array
$resultarray = @() 
#main loop
foreach ($computer in $computerlist) {
    
    $connection = $false
    $counter = 0
    #$computername = $computer.computer_name
    $computername = $computer.Name
    $computerdesc = $computer.Description
    Write-Progress -Activity 'Processing computer shares' -CurrentOperation $computername

    #temporary result to add to result array
    $result = New-Object PSObject

    #try quick to ping, otherwise give up
    if (Test-Connection -ComputerName $computername -Quiet -Count 1) {
        #use net view to see if there are any shares or printers
        $netview = net view \\$computername
        #parse net view output
        foreach ($line in $netview) {
            if ($line -match "Print") {
                #for every printer advance the counter
                $counter += 1
            }
        }
        if ($counter -gt 0) {
            #write-host "$computername has $counter printers shared"
            $result | Add-Member -MemberType NoteProperty -Name "Computer" -Value $computername
            $result | Add-Member -MemberType NoteProperty -Name "Description" -Value $computerdesc
            $result | Add-Member -MemberType NoteProperty -Name "Printer Count" -Value $counter
            $result | Add-Member -MemberType NoteProperty -Name "Shared Printers" -Value "yes"
        } else {
            #write-host "$computername has no printers shared"
            $result | Add-Member -MemberType NoteProperty -Name "Computer" -Value $computername
            $result | Add-Member -MemberType NoteProperty -Name "Description" -Value $computerdesc
            $result | Add-Member -MemberType NoteProperty -Name "Printer Count" -Value $counter
            $result | Add-Member -MemberType NoteProperty -Name "Shared Printers" -Value "no"
        }
    } else {
        #write-host "$computername is not reachable"
        $result | Add-Member -MemberType NoteProperty -Name "Computer" -Value $computername
        $result | Add-Member -MemberType NoteProperty -Name "Description" -Value $computerdesc
        $result | Add-Member -MemberType NoteProperty -Name "Printer Count" -Value "0"
        $result | Add-Member -MemberType NoteProperty -Name "Shared Printers" -Value "unreachable"
    }
    $resultarray += $result
}
#export to csv
$resultarray | export-csv $csvfilepath -NoTypeInformation
