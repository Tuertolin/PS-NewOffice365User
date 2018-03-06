<#
Sript to create a new user in Office 365 platform.
With this script you will be able to:
    1- Connect to an Office365 environment 
    2- Create a new user.
    3- Configure Region (8 available, you can change them)
    4- Selecting licenses
#>

Begin
{
    #Connect to the Office 365 platform
    Import-Module MSOnline
    $Cred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    Connect-MsolService -Credential $Cred
        
    ## FUNCTIONS ##
    function assign-OfficeGroup{  #Receives $UPN and $LOCATION and assign it to the corresponding Office Group.
        
        Param ($UPN,$LOCATION)

        switch ($LOCATION)
        {                
         "AU" { Add-UnifiedGroupLinks AUOffice –Links $UPN –LinkType Member; break}
         "PH" { Add-DistributionGroupMember -Identity MakatiOffice -Member $UPN ; break}
         "IN" { Add-DistributionGroupMember -Identity IndiaOffice -Member $UPN; break}
         "UK" { Add-DistributionGroupMember -Identity UKOffice -Member $UPN; break}
         "US" { Add-DistributionGroupMember -Identity DenverOffice -Member $UPN; break}
         "CA" { Add-DistributionGroupMember -Identity TorontoOffice -Member $UPN; break}
         "SG" { Add-DistributionGroupMember -Identity SingaporeOffice -Member $UPN; break}
         "NZ" { Add-DistributionGroupMember -Identity NZOffice -Member $UPN; break}
         default {break}
        }    
    }
}

Process
{ 
#Parameters
    $FirstName = Read-Host "Please enter user first name "
    $LastName = Read-Host "Please enter user last name "
    $TLD = Read-Host "Please enter your tld including the @ (eg. @yourdomain.com) "
    $Pass = $FirstName + $LastName.Substring(0,1) + '$123' ;
    $UPN = $FirstName + "." + $LastName + $TLD
    $Phone = Read-Host "Please enter phone number (Optional) "
    $JobTitle = Read-Host "Please enter Job Title (Optional) "
    
    Write-Host " 1 for NZ"
    Write-Host " 2 for AU"
    Write-Host " 3 for PH"
    Write-Host " 4 for IN"
    Write-Host " 5 for UK"
    Write-Host " 6 for US"
    Write-Host " 7 for CA"
    Write-Host " 8 for SG"
    $LOCATION = Read-Host "Please select the location for " $FirstName

    switch ($LOCATION)
        {
            1 {$LOCATION = "NZ"; break}
            2 {$LOCATION = "AU"; break}
            3 {$LOCATION = "PH"; break}
            4 {$LOCATION = "IN"; break}
            5 {$LOCATION = "UK"; break}
            6 {$LOCATION = "US"; break}
            7 {$LOCATION = "CA"; break}
            8 {$LOCATION = "SG"; break}
            default {$LOCATION = "AU"; break}
        }
    Write-Host "The location selected is: " $LOCATION
         
    $SKU = Get-MSOLAccountSKU
        if($SKU.count -eq 1)
        {
            Write-Host "Only one SKU in this tenant."
        }
        elseif($SKU.count -gt 1)
        {
            $i = 0
            do
            {
                Write-Host "$i,$($SKU[$i].AccountSkuId)"
                $i++
            }while($i -le $SKU.count-1)
            [int]$userChoice = Read-Host "Please select the License for " $FirstName
            $License = $SKU[$userChoice].AccountSkuId
        }

    Write-Host
    Write-Host "#######################"
    Write-Host "The user will be created as:"
    Write-Host "First name: " $FirstName
    Write-Host "Last name: " $LastName
    Write-Host "UPN Name: " $UPN
    Write-Host "Pass: " $Pass
    Write-Host "Location: " $LOCATION
    Write-Host "License: " $License
    Write-Host
    $Confirm = Read-Host "Do you confirm that: (Y,y or N,n)"
    $Confirm = $Confirm.ToUpper()

    if ($Confirm -eq "Y")
        {
            Write-Host "The user is creating..."
            <#Script: First Will create the MSOL user.
            Assigning: License, Group Membership, PhoneNumber, JobTitle #>
            
            New-MsolUser -UserPrincipalName $UPN -FirstName $FirstName -LastName $LastName `
                -DisplayName "$($FirstName) $($LastName)" `
                –Password $Pass -ForceChangePassword $False -PasswordNeverExpires $True –UsageLocation $LOCATION      
      
            Set-MsolUserLicense –UserPrincipalName $UPN -AddLicenses $License
            assign-OfficeGroup $UPN $LOCATION
            Set-msoluser -UserPrincipalName $UPN -PhoneNumber $Phone
            Set-msoluser -UserPrincipalName $UPN -Title $JobTitle

        }
        
        else
            {Write-Host "User Creation cancelled."}
}

End
{}
