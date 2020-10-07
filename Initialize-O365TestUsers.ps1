function New-SecurePassPhrase {
    $TextInfo = (Get-Culture).TextInfo

    $jsonFiles = Get-ChildItem (Join-Path (Get-Location) "\JSON")
    $bigList = @();
    foreach($file in $jsonFiles.FullName) {
        $bigList += Get-Content $file | ConvertFrom-Json
    }
    $bigList = $bigList | Where-Object {$_.PartOfSpeech}
    $adjList = $bigList | Where-Object {$_.PartOfSpeech -eq "adjective" -and $_.word.length -ge 3 -and $_.word.length -le 6}
    $nounList = $bigList | Where-Object {$_.PartOfSpeech -eq "noun" -and $_.word.length -ge 3 -and $_.word.length -le 6 }

    $adj = $TextInfo.toTitleCase($adjList[(Get-Random -Minimum 0 -Maximum $adjList.length)].word)
    $noun = $TextInfo.toTitleCase($nounList[(Get-Random -Minimum 0 -Maximum $adjList.length)].word)
    $randNumber = "{0:000}" -f (Get-Random -Minimum "9" -Maximum "999")

    return "{0}{1}{2}" -f $adj, $noun, $randNumber
}

function Initialize-O365TestUsers {
    param(
    [int]$Count = 250,
    [Parameter(Mandatory=$True)][String]$CountryCode,
    [System.Management.Automation.PSCredential]$Credential
    )

    begin {
        if($Count -gt 250) { $Count = 250 }
        try {
            if(-not $Credential) { 
                $Credential = Get-Credential -Message "Enter your name/password for your Azure lab tenancy"
            }
            if(-not (Get-Module AzureAd)) { Install-Module AzureAD -Confirm:$False -AllowClobber }

            Import-Module AzureAD

            Connect-AzureAD -Credential $Credential
        }
        catch {
            Throw "Error initialising script. Make sure you are able to install the AzureAD PowerShell module (may require launching PowerShell with Administrator privileges)."
        }
    }
    
    process {
        try {
            $DomainSuffix = (Get-AzureAdDomain | Where-Object {$_.IsDefault -eq $True}).Name
        }
        catch {
            throw "You logged in as your external e-mail address, which cannot poll the AzureAD Domain Name!"
        }

        do {
            try {
                if(-not $Count) { $Count = 250 }
                $users = Invoke-RestMethod -Method Get -Uri "https://randomuser.me/api/1.3/?results=$($Count)&nat=AU" -ErrorAction SilentlyContinue
            }
            catch {
                Write-Information -MessageData "Failed to invoke API resource correctly, trying again in 3 seconds.."
                Start-Sleep -Seconds 3
            }
        } while ($users.results.length -ne $Count)

        $Jobs = Import-Csv -Path "JobTitles.csv"

        foreach($user in $users.results) { 
            $Index = 1..50 | Get-Random
            $Department = $Jobs.Department[$Index]
            $Title = $Jobs.JobTitle[$Index]

            $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
            $Password = New-SecurePassPhrase
            $PasswordProfile.Password = $Password
            $PasswordProfile.EnforceChangePasswordPolicy = $false
            $PasswordProfile.ForceChangePasswordNextLogin = $false

            $EmsPremiumSkuId = (Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -eq "EMSPREMIUM"}).SkuId
            $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $License.SkuId = $EmsPremiumSkuId
            
            $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToAssign.AddLicenses = $License

            New-AzureADUser -UserPrincipalName "$($user.name.first).$($user.name.last)@$($DomainSuffix)" -City $user.location.City -DisplayName "$($user.name.first) $($user.name.last)" -GivenName $user.name.first -Surname $user.name.last -Country $user.location.country -State $user.location.State -JobTitle $Title -Department $Department -UsageLocation $CountryCode -PasswordProfile $PasswordProfile -PostalCode $user.location.postcode -ShowInAddressList:$True -AccountEnabled:$True -MailNickName "$($user.name.first).$($user.name.last)"
            Set-AzureADUserLicense -ObjectId "$($user.name.first).$($user.name.last)@$($DomainSuffix)" -AssignedLicenses $LicensesToAssign

            $UserDetails = [PSCustomObject]@{
                Name = "($user.name.first) $($user.name.last)"
                Email = "($user.name.first).$($user.name.last)@$($DomainSuffix)"
                Password = $Password
            }
            $UserDetails
        }

        Start-Sleep -Seconds 60

        foreach($user in $users.results) {
            try {
                # Create a directory for photos if it doesn't exist
                if(-not (Test-Path (Join-Path (Get-Location) "Photos"))) { New-Item -ItemType Directory -Name (Join-Path (Get-Location) "Photos") }
    
                $ObjectId = (Get-AzureAdUser -ObjectId "$($user.name.first).$($user.name.last)@$($DomainSuffix)").ObjectId
                [Byte[]]$result = (Invoke-WebRequest -Uri $user.picture.large).Content
                [System.IO.File]::WriteAllBytes((Join-Path (Get-Location) "Photos\$(Split-Path $user.picture.large -Leaf)"), $result) | Out-Null
                Set-AzureADUserThumbnailPhoto -ObjectId $ObjectId -FilePath (Join-Path (Get-Location) "Photos\$(Split-Path $user.picture.large -Leaf)")
            }
            catch {
                Write-Error -Message "Failed to upload photo for $($user.name.first) $($user.name.last)`nObjectId: $ObjectId"
            }
        }
    }
    end {
        Disconnect-AzureAD
    }
}

Start-Transcript -Path "O365UserCreds.txt"
$Region = (Invoke-RestMethod -Method Get -Uri "https://restcountries.eu/rest/v2/all") | Select-Object Name, @{n="CountryCode";e={$_.alpha2Code}} | ` Out-GridView -Title "Select the country you want to create all the accounts in" -OutputMode Single
$Count = Read-Host -Prompt "How many users do you want to create? (max. 250 for EMSPREMIUM license)."
Initialize-O365TestUsers -Count $Count -CountryCode $Region.CountryCode
Stop-Transcript