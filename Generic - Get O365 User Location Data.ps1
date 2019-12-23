function Merge-CSVFiles { 
[cmdletbinding()] 
param( 
    [string[]]$CSVFiles, 
    [string]$OutputFile = "c:\merged.csv" 
) 
$Output = @(); 
foreach($CSV in $CSVFiles) { 
    if(Test-Path $CSV) { 
         
        $FileName = [System.IO.Path]::GetFileName($CSV) 
        $temp = Import-CSV -Path $CSV | select *, @{Expression={$FileName};Label="FileName"} 
        $Output += $temp 
 
    } else { 
        Write-Warning "$CSV : No such file found" 
    } 
 
} 
$Output | Export-Csv -Path $OutputFile -NoTypeInformation 
Write-Output "$OutputFile successfully created" 
 
} 

del /temp/*.csv


Function Connect-EXOnline {
    $credentials = Get-Credential -Credential $credential
     
    $Session = New-PSSession  -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber
}

Connect-EXOnline
  
 
$startDate = (Get-Date).AddDays(-10)
$endDate = (Get-Date)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox  #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData0-10Days.csv -Append -NoTypeInformation
    }
}

 
$startDate = (Get-Date).AddDays(-20)
$endDate = (Get-Date).AddDays(-11)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData11-20Days.csv -Append -NoTypeInformation
    }
}

 
$startDate = (Get-Date).AddDays(-30)
$endDate = (Get-Date).AddDays(-21)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData21-30Days.csv -Append -NoTypeInformation
    }
}

$startDate = (Get-Date).AddDays(-40)
$endDate = (Get-Date).AddDays(-31)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData31-40Days.csv -Append -NoTypeInformation
    }
}

$startDate = (Get-Date).AddDays(-50)
$endDate = (Get-Date).AddDays(-41)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData41-50Days.csv -Append -NoTypeInformation
    }
}

$startDate = (Get-Date).AddDays(-60)
$endDate = (Get-Date).AddDays(-51)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData51-60Days.csv -Append -NoTypeInformation
    }
}

$startDate = (Get-Date).AddDays(-70)
$endDate = (Get-Date).AddDays(-61)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData61-70Days.csv -Append -NoTypeInformation
    }
}

$startDate = (Get-Date).AddDays(-80)
$endDate = (Get-Date).AddDays(-71)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData71-80Days.csv -Append -NoTypeInformation
    }
}

$startDate = (Get-Date).AddDays(-90)
$endDate = (Get-Date).AddDays(-81)
$Logs = @()
    Write-Host "Retrieving logs" -ForegroundColor Red
    do {
        $logs += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId "UALSearch" -ResultSize 5000 -StartDate $startDate -EndDate $endDate -Operations UserLoggedIn, New-InboxRule, Set-Mailbox #-SessionId "$($customer.name)"
        Write-Host "Retrieved $($logs.count) logs" -ForegroundColor Yellow
    }while ($Logs.count % 5000 -eq 0 -and $logs.count -ne 0)
    Write-Host "Finished Retrieving logs" -ForegroundColor Green
 
$userIds = $logs.userIds | Sort-Object -Unique
 
foreach ($userId in $userIds) {
 
    $ips = @()
    Write-Host "Getting logon IPs for $userId"
    $searchResult = ($logs | Where-Object {$_.userIds -contains $userId}).auditdata | ConvertFrom-Json -ErrorAction SilentlyContinue
    Write-Host "$userId has $($searchResult.count) logs" -ForegroundColor Green
 
    $ips = $searchResult.clientip | Sort-Object -Unique
    Write-Host "Found $($ips.count) unique IP addresses for $userId"
    foreach ($ip in $ips) {
        Write-Host "Checking $ip" -ForegroundColor Yellow
        $mergedObject = @{}
        $singleResult = $searchResult | Where-Object {$_.clientip -contains $ip} | Select-Object -First 1
        Start-sleep -m 400
        $ipresult = Invoke-restmethod -method get -uri http://ip-api.com/json/$ip
        $UserAgent = $singleResult.extendedproperties.value[0]
        Write-Host "Country: $($ipResult.country) UserAgent: $UserAgent"
        $singleResultProperties = $singleResult | Get-Member -MemberType NoteProperty
        foreach ($property in $singleResultProperties) {
            if ($property.Definition -match "object") {
                $string = $singleResult.($property.Name) | ConvertTo-Json -Depth 10
                $mergedObject | Add-Member -Name $property.Name -Value $string -MemberType NoteProperty    
            }
            else {$mergedObject | Add-Member -Name $property.Name -Value $singleResult.($property.Name) -MemberType NoteProperty}          
        }
        $property = $null
        $ipProperties = $ipresult | get-member -MemberType NoteProperty
 
        foreach ($property in $ipProperties) {
            $mergedObject | Add-Member -Name $property.Name -Value $ipresult.($property.Name) -MemberType NoteProperty
        }
        $mergedObject | Select-Object UserId, Operation, CreationTime, @{Name = "UserAgent"; Expression = {$UserAgent}}, Query, ISP, City, RegionName, Country  | export-csv C:\temp\UserLocationData81-90Days.csv -Append -NoTypeInformation
    }
}




Get-PSSession | Remove-PSSession
Merge-CSVFiles -CSVFiles C:\temp\UserLocationData0-10Days.csv,C:\temp\UserLocationData11-20Days.csv,C:\temp\UserLocationData21-30Days.csv,C:\temp\UserLocationData31-40Days.csv,C:\temp\UserLocationData41-50Days.csv,C:\temp\UserLocationData51-60Days.csv,C:\temp\UserLocationData61-70Days.csv,C:\temp\UserLocationData71-80Days.csv,C:\temp\UserLocationData81-90Days.csv C:\temp\O365UserLocationData.csv


