param(
    [switch]$Test,
    [switch]$Set,
    [switch]$Restore,
    [ValidateNotNullOrEmpty()]$ExportCSV,
    [ValidateNotNullOrEmpty()]$ErrorTabCSV,
    [ValidateNotNullOrEmpty()]$OldDirectionName,
    [ValidateNotNullOrEmpty()]$NewDirectionName,
    $DirectionOU,
    $TestFile,
    $SetFileToRestore,
    $dc = [String]((Get-ADDomainController -DomainName <domain> -Discover | Select-Object -First 1).hostname)
)

<#
.DISCLAIMER
    This script has been created for one customer and based on his infrastructure and his engineer rules. It means that it is not generic

.SYNOPSIS
    This scripts is used in a project which is renaming objects that have the direction in their attributes. This one is focused on the diffusion lists. 
    
.DESCRIPTION
    This script allows to renamme the followings attributes:
        - UserPrincipalName
        - Mail Adress
        - DisplayName
        - Department
        - ExtensionAttribute4
        - SamAccountName
        - Name
    And the Mail Adress Policy

    This script got 3 Steps:
        - 1st step: TEST    : Simulate the new values and check if the attributes are not already used. No parameters are set in this step, it's just a simulation
        - 2nd step: SET     : Based on the CSV generated while the 1st step (TEST), it will set the attributes in function of this one.
        - 3rd step: RESTORE : Based on the CSV generated while the 1st step (SET), this step allows you to do a rollback.

.EXAMPLE
    1. TEST Step
        $Test = & '.\Script LD Directions.ps1' -Test -OldDirectionName "DirectionX" -NewDirectionName "DirectionZ" -DirectionOU "Name de l'OU" -ExportCSV "C:\Users\admin-msadekha\Documents\result\LD-Modifications.csv" -ErrorTabCSV "C:\Users\admin-msadekha\Documents\logs\ERROR-LD-Modifications.csv"


    2. SET Step
        $Set = & '.\Script LD Directions.ps1' -Set -TestFile "CSV généré par la phase de TEST ci-dessus" -NewDirectionName "DirectionZ" -ExportCSV "C:\Users\admin-msadekha\Documents\result\LD-Modifications.csv" -ErrorTabCSV "C:\Users\admin-msadekha\Documents\logs\ERROR-LD-Modifications.csv"


    3. RESTORE Step
        $Restore = & '.\Script LD Directions.ps1' -Restore -SetFileToRestore "CSV généré par la phase de SET ci-dessus" -OldDirectionName "DirectionX" -ExportCSV "C:\Users\admin-msadekha\Documents\result\LD-Modifications.csv" -ErrorTabCSV "C:\Users\admin-msadekha\Documents\logs\ERROR-LD-Modifications.csv"
        
.NOTES
    Author: Hakim M'SADEK

    Date: 08 Avril 2022

.VERSION
    Version : 1.0
    
#>

#Function wich collect the princpal and secondary mail adresses
Function Get-PrincipalAndSecondary_Addresses
{
    param(
        [Parameter(Mandatory)]$ProxyAddresses
    )
    $TabProperties = @{
        PrincipalAddress = $null
        SecondaryAddress = $null
    }

    $Tab = New-Object PSObject -Property $TabProperties
    ForEach ($Address in ($LD.proxyaddresses -split ",")){
        if($Address -clike "SMTP*"){
            if(!$PrincipalAdress){
                $PrincipalAdress = $Address.replace("SMTP:","")
            }
            else{
                $PrincipalAdress = $PrincipalAdress + ","+ $Address.replace("SMTP:","")
            }
        }
        elseif($Address -clike "smtp*"){
            if(!$SecondaryAdress){
                $SecondaryAdress = $Address.replace("smtp:","")
            }
            else{
                $SecondaryAdress = $SecondaryAdress + "," + $Address.replace("smtp:","")
            }
        }
    }
    $Tab.PrincipalAddress = $PrincipalAdress
    $Tab.SecondaryAddress = $SecondaryAdress

    Return $Tab
}

#Function wich set a line in the error tab
function Set-ErrorTabline 
{
    param(
        [ValidateNotNullOrEmpty()]$LD,
        [ValidateNotNullOrEmpty()]$ErrorType, #ERROR #CRITICALERROR
        [ValidateNotNullOrEmpty()]$ErrorDescription,
        [ValidateNotNullOrEmpty()]$ErrorMessage
    )
    $ErrorTabproperties = @{
        Time = $null
        DisplayName = $null
        ProxyAddresses = $null
        ObjectGUID = $null
        Error = $null
        ErrorType = $null
    }
    $curTime = Get-Date -Format "dd'/'MM'/'yyyy hh':'mm':'ss"
    write-host "[$ErrorType] $ErrorDescription | ObjectGUID : $($LD.ObjectGUID) | OldDisplayName : $($LD.DisplayName)" -ForegroundColor red

    $TablineError = New-Object PSObject -Property $ErrorTabproperties
    $TablineError.ErrorType = $ErrorType
    $TablineError.Time = $curTime
    $TablineError.DisplayName = $DisplayName
    $TablineError.ProxyAddresses = $LD.ProxyAddresses -join ","
    $TablineError.ObjectGUID = $LD.ObjectGuid
    $TablineError.Error = "[$ErrorDescription] $ErrorMessage"

    Return $TablineError
}

function Test-PrincipalAddress
{
    param(
        [ValidateNotNullOrEmpty()]$LD,
        $Tabline,
        $PrincipalAddressLow,
        $SecondaryAddressLow,
        [ValidateNotNullOrEmpty()]$OldDirectionLow,
        [ValidateNotNullOrEmpty()]$NewDirectionLow,
        [ValidateNotNullOrEmpty()]$ErrorTabCSV
    )
    $OldDirectionLength = $OldDirectionLow.Length
    #ERROR if no mailbox
    if (!$(Get-Recipient $PrincipalAddressLow -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Where-Object Guid -eq $LD.ObjectGuid)){
        Set-ErrorTabline -LD $LD -ErrorDescription "Do not have any Mailbox" -ErrorType "ERROR" -ErrorMessage "Do not have any Mailbox" | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
        continue
    }

    #ERROR if got multiple principal adress
    if(($PrincipalAddressLow -split ",").Length -ne 1)
    {
        Set-ErrorTabline -LD $LD -ErrorDescription "PrincipalAddress contain more than 1 address" -ErrorType "ERROR" -ErrorMessage "PrincipalAddress contain more than 1 address" | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
        continue
    }

    #
    if((Get-Recipient -Identity $PrincipalAddressLow).EmailAddressPolicyEnabled -eq $True){
        $Tabline."MailRule Set" = "Yes"
    }
    else{
        $Tabline."MailRule Set" = "No"
    }

    #Construction of the new principal mail
    $Tabline.OldPrincipalAddress = $PrincipalAddressLow
    if ($PrincipalAddressLow -like "*$OldDirectionLow*"){
        $Tabline.NewPrincipalAddress = $($NewDirectionLow + $PrincipalAddressLow.Substring($OldDirectionLength))
    }
    #If the adress doesn't contain an old direction name, it doesn't modify the attribute and get out from this function
    else{
        $tabline.NewPrincipalAddress = $PrincipalAddressLow
        $Tabline.NewSecondaryAddress = $SecondaryAddressLow
        $Tabline.OldSecondaryAddress = $SecondaryAddressLow
        $Tabline."NewPrincipalAddress set" = "Yes"
        $Tabline."NewSecondaryAddress set" = "Yes"

        Return $Tabline
    }
    #Test if the simulated principal mail adress already exists
    $TestPrincipalAddress = Get-Recipient $Tabline.NewPrincipalAddress -ErrorAction SilentlyContinue | Where-Object Guid -ne $LD.ObjectGuid
    if($TestPrincipalAddress){
        Set-ErrorTabline -LD $LD -ErrorDescription "NewPrincipalAddress ($($Tabline.NewPrincipalAddress)) already exists" -ErrorType "ERROR" -ErrorMessage "NewPrincipalAddress ($($Tabline.NewPrincipalAddress)) already exists" | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
        continue
    }
    #Test if the simulated principal mail adress is already set
    if ($Tabline.NewPrincipalAddress -eq $PrincipalAddressLow)
    {
        $Tabline."NewPrincipalAddress set" = "Yes"
        Set-ErrorTabline -LD $LD -ErrorType "ERROR" -ErrorDescription "Cannot set the New Secondary address because the New Principal Address is already Set" -ErrorMessage "Cannot set the New Secondary address because the New Principal Address is already Set" | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
    }
    else{
        $Tabline."NewPrincipalAddress set" = "No"
    }

    Return $Tabline

}

function Test-SecondaryAddress
{
    param(
        [ValidateNotNullOrEmpty()]$LD,
        $Tabline,
        [ValidateNotNullOrEmpty()]$PrincipalAddressLow,
        $SecondaryAddressLow,
        [ValidateNotNullOrEmpty()]$OldDirectionLow,
        [ValidateNotNullOrEmpty()]$NewDirectionLow
    )
    if ($Tabline."NewPrincipalAddress set" -like "No")
    {
        
        ############ SECONDARYADDRESS SECTION ############
        $Tabline.OldSecondaryAddress = $SecondaryAddressLow
    
        if ($SecondaryAddressLow){
            $Tabline.NewSecondaryAddress = $PrincipalAddressLow + "," + $SecondaryAddressLow
        }
        else{
            $Tabline.NewSecondaryAddress = $PrincipalAddressLow
        }
        #Test si le NewSecondaryAddress est dï¿½jï¿½ set
        if ($Tabline.OldSecondaryAddress){
            if (($Tabline.OldSecondaryAddress -split ",") -notcontains $PrincipalAddressLow){
                $Tabline."NewSecondaryAddress set" = "No"
            }
            else{
                $Tabline."NewSecondaryAddress set" = "Yes"
            }
        }
        else{
            $Tabline."NewSecondaryAddress set" = "No"
        }
    }

    Return $Tabline
}

Function Test-DirectionAttribute
{
    param(
        $Tabline,
        [ValidateNotNullOrEmpty()]$Attribute,
        [ValidateNotNullOrEmpty()]$LD,
        [ValidateNotNullOrEmpty()]$NewDirectionName
    )
    $Tabline."Old$Attribute" = $LD.$Attribute
    $Tabline."New$Attribute" = $NewDirectionName
    if ($Tabline."Old$Attribute" -eq $NewDirectionName){
        $Tabline."New$Attribute set" = "Yes"
    }
    else{
        $Tabline."New$Attribute set" = "No"
    }

    Return $Tabline
}

Function Set-Addresses
{
    param(
        [ValidateNotNullOrEmpty()]$LD,
        $Tabline,
        [ValidateNotNullOrEmpty()]$ErrorTabCSV,
        [ValidateSet("Old","New")]$OldOrNew,
        [ValidateNotNullOrEmpty()]$dc

    )

    if($OldOrNew -eq "Old"){
        $YesOrNo = "Yes"
    }
    elseif($OldOrNew -eq "New"){
        $YesOrNo = "No"
    }

    if ($LD."NewSecondaryAddress set" -like $YesOrNo -and $LD."NewPrincipalAddress set" -like $YesOrNo)
    {
        try{
            Set-DistributionGroup -Identity $LD.Objectguid -Alias (($LD."$($OldOrNew)PrincipalAddress") -Split "@")[0] -DomainController $dc -ErrorAction Stop -WarningAction Silentlycontinue
            if($LD."MailRule Set" -like "No"){
                Set-DistributionGroup -Identity $LD.Objectguid -EmailAddressPolicyEnabled $True -DomainController $dc -ErrorAction Stop -WarningAction Silentlycontinue
            }
            $Tabline."MailRule Set" = "Yes"
        }
        catch{
            Set-ErrorTabline -BalObj $LD -ErrorType "ERROR" -ErrorDescription "Set $OldOrNew Principal Address" -ErrorMessage $_.exception.message | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
            continue
        }
        $Tabline.NewPrincipalAddress = $LD.NewPrincipalAddress
        $Tabline.NewSecondaryAddress = $LD.NewSecondaryAddress
        $tabline.OldPrincipalAddress = $LD.OldPrincipalAddress
        $tabline.OldSecondaryAddress = $LD.OldSecondaryAddress
        $Tabline."$($OldOrNew)PrincipalAddress set" = "Yes"
        $Tabline."$($OldOrNew)SecondaryAddress set" = "Yes"

        #Lopp which wait for the modifications to be apllied in AD
        while ((Get-ADGroup -Server <domain> -Identity $LD.ObjectGUID -Properties ProxyAddresses).ProxyAddresses -notcontains "SMTP:$($LD."$($OldOrNew)PrincipalAddress")")
        {
            Start-Sleep 1
        }
    }
    else{
        $Tabline.NewPrincipalAddress = $LD.NewPrincipalAddress
        $Tabline.NewSecondaryAddress = $LD.NewSecondaryAddress
        $tabline.OldPrincipalAddress = $LD.OldPrincipalAddress
        $tabline.OldSecondaryAddress = $LD.OldSecondaryAddress
        $Tabline."$($OldOrNew)PrincipalAddress set" = "Yes"
        $Tabline."$($OldOrNew)SecondaryAddress set" = "Yes"
    }
    Return $Tabline
}

Function Set-AdObjAttribute
{
    param(
        [ValidateNotNullOrEmpty()]$Adobj,
        [ValidateNotNullOrEmpty()]$Obj,
        [ValidateNotNullOrEmpty()]$Attribute,
        $Value,
        [ValidateSet("Old","New")]$OldOrNew
    )
    if ($OldOrNew -eq "New"){
        $YesorNo = "No"
    }
    elseif($OldOrNew -eq "Old"){
        $YesorNo = "Yes"
        if (!$Obj."Old$Attribute"){
            $ADObj.$Attribute = $Null
            Return $Adobj
        }
    }


    if ($Obj."New$Attribute set" -eq $YesorNo){
        $ADObj.$Attribute = $Value
    }

    Return $Adobj
}

Function Rename-LD-Name
{
     param(
        [ValidateNotNullOrEmpty()]$LD,
        [ValidateNotNullOrEmpty()]$ErrorTabCSV,
        $Tabline,
        [ValidateSet("Old","New")]$OldOrNew
    )
    if ($OldOrNew -eq "New"){
        $YesorNo = "No"
    }
    elseif ($OldOrNew -eq "Old"){
        $YesorNo = "Yes"
    }

    if ($LD."NewName set" -eq $YesorNo){
        try{
            Get-ADObject -Server <domain> -Identity $LD.ObjectGUID | Rename-ADObject -NewName $LD."$($OldOrNew)Name"
        }
        catch{
            Set-ErrorTabline -LD $LD -ErrorType "ERROR" -ErrorDescription "Set $OldOrNew Name" -ErrorMessage $_.exception.message  | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
            continue
        }
    }
    $Tabline."NewName" = $LD."NewName"
    $Tabline."OldName" = $LD."OldName"
    $Tabline."$($OldOrNew)Name set" = "Yes"
    Return $Tabline
}

Function Set-SpecialAttribute
{
    param(
        [ValidateNotNullOrEmpty()]$LD,
        $Tabline,
        [ValidateNotNullOrEmpty()]$Attribute,
        [ValidateSet("Old","New")]$OldOrNew
    )
    if ($OldOrNew -eq "New"){
        $YesorNo = "No"
    }
    elseif($OldOrNew -eq "Old"){
        $YesorNo = "Yes"
    }

    if ($LD."New$Attribute set" -eq $YesorNo){
        try{
            (Get-ADGroup -Server <domain> -Identity $LD.ObjectGUID) | Set-ADGroup -Replace @{$Attribute=$($LD."$($OldOrNew)$Attribute")} -ErrorAction Stop
        }
        catch{
            if ($Attribute -eq "SamAccountName"){
                Set-ErrorTabline -LD $LD -ErrorType "ERROR" -ErrorDescription "Set $OldOrNew$Attribute (20Char max)" -ErrorMessage $_.exception.message  | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
                continue
            }

            Set-ErrorTabline -LD $LD -ErrorType "ERROR" -ErrorDescription "Set $OldOrNew$Attribute" -ErrorMessage $_.exception.message  | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
            continue
        }
        
        #Lopp which wait for the modifications to be apllied in AD
        while ((Get-ADGroup -Server <domain> -Identity $LD.ObjectGUID -Properties $Attribute).$Attribute -notlike $LD."$($OldOrNew)$Attribute")
        {
            Start-Sleep 1
        }
    }
    $Tabline."Old$Attribute" = $LD."Old$Attribute"
    $Tabline."New$Attribute" = $LD."New$Attribute"
    $Tabline."$($OldOrNew)$Attribute set" = "Yes"

    Return $Tabline
}

Function Test-AttributeWithCheck
{

    param(
        $Tabline,
        [ValidateNotNullOrEmpty()]$LD,
        [ValidateNotNullOrEmpty()]$Attribute,
        [ValidateNotNullOrEmpty()]$OldDirectionName,
        [ValidateNotNullOrEmpty()]$NewDirectionName,
        [ValidateNotNullOrEmpty()]$ErrorTabCSV
    )

    if($LD.$Attribute -like "*@*"){
        $OldDirectionName = $OldDirectionName.ToLower()
        $NewDirectionName = $NewDirectionName.ToLower()
    }

    $Tabline."Old$Attribute" = $LD.$Attribute
    if($LD.$Attribute -like "*$OldDirectionName*"){
        $Tabline."New$Attribute" = $LD.$Attribute.Replace($OldDirectionName,$NewDirectionName)
    }
    else{ #If the attributes doesn't start with the new or the old direction, nothing is done
        $Tabline."New$Attribute" = $LD.$Attribute
        $Tabline."Old$Attribute" = $LD.$Attribute
        $Tabline."New$Attribute set" = "Yes"

        Return $Tabline
    }
    
    #Test if the attribute is already used
    $TestAttribute = Get-ADGroup -Server <domain> -LDAPFilter "($Attribute=$($Tabline."New$Attribute"))" | Where-Object ObjectGUID -ne $LD.ObjectGuid -ErrorAction SilentlyContinue
    if($TestAttribute)
    {
        Set-ErrorTabline -LD $LD -ErrorDescription "New$Attribute ($($Tabline."New$Attribute")) already exists" -ErrorType "ERROR" -ErrorMessage "New$Attribute ($($Tabline."New$Attribute")) already exists" | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
        continue
    }

    if ($Tabline."New$Attribute" -eq $LD.$Attribute)
    {
        $Tabline."New$Attribute set" = "Yes"
    }
    else {
        $Tabline."New$Attribute set" = "No"
    }
    Return $Tabline

}


#___________________________________________________________________________________________________________________________________________________________________________________________________#




Function Get-LDs
{
    param(
        [Parameter(Mandatory)]$OldDirectionName,
        [Parameter(Mandatory)]$DirectionOU
    )
    $OU = <Distribution list distinguished name with the $directionOU name>
    Return (Get-ADGroup -Server <domain> -SearchBase $OU -filter "(GroupCategory -eq 'Distribution')" -Properties *)
}

Function Test-LD-Modifications
{
    Param(
        [Parameter(Mandatory)]$LD,
        [Parameter(Mandatory)]$OldDirectionName,
        [Parameter(Mandatory)]$NewDirectionName,
        [Parameter(Mandatory)]$ExportCSV,
        [Parameter(Mandatory)]$ErrorTabCSV
    )
    #Define the CSV Name
    $OldExportCSVFilename = $ExportCSV.split("\")[-1]
    $NewExportCSVFilename = $OldExportCSVFilename.replace($OldExportCSVFilename, "(Test)$OldExportCSVFilename")

    $OldErrorTabCSVFileName = $ErrorTabCSV.split("\")[-1]
    $NewErrorTabCSVFileName = $OldErrorTabCSVFileName.replace($OldErrorTabCSVFileName, "(Test)$OldErrorTabCSVFileName")

    $ExportCSV = $ExportCSV.replace($OldExportCSVFilename, $NewExportCSVFilename)
    $ErrorTabCSV = $ErrorTabCSV.replace($OldErrorTabCSVFileName,$NewErrorTabCSVFileName)


    $OldDirectionName = $OldDirectionName.ToUpper()
    $NewDirectionName = $NewDirectionName.ToUpper()
    $NewDirectionLow = $NewDirectionName.ToLower()
    $OldDirectionLow = $OldDirectionName.ToLower()

    #Properties of the result table
    $TabProperties = @{
        OldDisplayName = $null
        NewDisplayName = $null
        OldName =$null
        NewName = $null
        ObjectGuid = $null
        OldPrincipalAddress = $null
        NewPrincipalAddress = $null
        OldSecondaryAddress = $null
        NewSecondaryAddress = $null
        CanonicalName = $null
        OldDepartment = $null
        NewDepartment = $null
        OldExtensionAttribute4 = $null
        NewExtensionAttribute4 = $null
        OldSamAccountName = $null
        NewSamAccountName = $null
        "NewPrincipalAddress set" = $null
        "NewSecondaryAddress set" = $null
        "NewDisplayName set" = $null
        "NewDepartment set" = $null
        "NewExtensionAttribute4 set" = $null
        "NewSamAccountName set" = $null
        "NewName set" = $null
        "MailRule Set"= $null
    }

    $ProxyAddresses = Get-PrincipalAndSecondary_Addresses -ProxyAddresses $LD.ProxyAddresses
    if($ProxyAddresses.PrincipalAddress){$PrincipalAddressLow = $($ProxyAddresses.PrincipalAddress).ToLower()}else{$PrincipalAddressLow=$null}
    if($ProxyAddresses.SecondaryAddress){$SecondaryAddressLow = $($ProxyAddresses.SecondaryAddress).ToLower()}else{$SecondaryAddressLow=$null}



    ################### ----------- DEBUT ----------- ###################
    
    
    
    $Tabline = New-Object PSObject -Property $TabProperties


    ############ DISPLAYNAME SECTION ############
    $Tabline = Test-AttributeWithCheck -Tabline $Tabline -LD $LD -Attribute "DisplayName" -OldDirectionName $OldDirectionName -NewDirectionName $NewDirectionName -ErrorTabCSV $ErrorTabCSV


    ############ PRINCIPAL ADDRESS SECTION ############
    $Tabline = Test-PrincipalAddress -LD $LD -Tabline $Tabline -PrincipalAddressLow $PrincipalAddressLow -SecondaryAddressLow $SecondaryAddressLow -OldDirectionLow $OldDirectionLow -NewDirectionLow $NewDirectionLow -ErrorTabCSV $ErrorTabCSV


    ############ SECONDARY ADDRESS SECTION ############
    $Tabline = Test-SecondaryAddress -LD $LD -Tabline $Tabline -PrincipalAddressLow $PrincipalAddressLow -SecondaryAddressLow $SecondaryAddressLow -OldDirectionLow $OldDirectionLow -NewDirectionLow $NewDirectionLow


    ############ SAMACCOUNTNAME SECTION ###########
    $Tabline = Test-AttributeWithCheck -Tabline $Tabline -LD $LD -Attribute "sAMAccountName" -OldDirectionName $OldDirectionName -NewDirectionName $NewDirectionName -ErrorTabCSV $ErrorTabCSV


    ############ NAME SECTION ###########
    $Tabline = Test-AttributeWithCheck -Tabline $Tabline -LD $LD -Attribute "Name" -OldDirectionName $OldDirectionName -NewDirectionName $NewDirectionName -ErrorTabCSV $ErrorTabCSV


    ############ DEPARTMENT SECTION ###########
    $Tabline = Test-DirectionAttribute -Tabline $Tabline -LD $LD -NewDirectionName $NewDirectionName -Attribute "Department"


    ############ EXTENSIONATTRIBUTE4 SECTION ############
    $Tabline = Test-DirectionAttribute -Tabline $Tabline -LD $LD -NewDirectionName $NewDirectionName -Attribute "ExtensionAttribute4"

    
    ################### ----------- Fin ----------- ###################


    $Tabline.CanonicalName = $LD.CanonicalName
    $Tabline.ObjectGuid = $LD.ObjectGuid.guid.ToString()

    $Tabline | Select-Object OldDisplayName, NewDisplayName, "NewDisplayName set", OldPrincipalAddress, NewPrincipalAddress, "NewPrincipalAddress set", OldSecondaryAddress, NewSecondaryAddress, "NewSecondaryAddress set", OldSamAccountName, NewSamAccountName, "NewSamAccountName set", OldName, NewName, "NewName set", OldDepartment, NewDepartment, "NewDepartment set", OldExtensionAttribute4, NewExtensionAttribute4, "NewExtensionAttribute4 set", "MailRule Set", ObjectGuid, CanonicalName | Export-csv -Path $ExportCSV -Delimiter ";" -Encoding Utf8 -NoTypeInformation -Append

    Return $Tabline | Select-Object OldDisplayName, NewDisplayName, "NewDisplayName set", OldPrincipalAddress, NewPrincipalAddress, "NewPrincipalAddress set", OldSecondaryAddress, NewSecondaryAddress, "NewSecondaryAddress set", OldSamAccountName, NewSamAccountName, "NewSamAccountName set", OldName, NewName, "NewName set", OldDepartment, NewDepartment, "NewDepartment set", OldExtensionAttribute4, NewExtensionAttribute4, "NewExtensionAttribute4 set", "MailRule Set", ObjectGuid, CanonicalName
}

Function Set-LDs-Modifications
{
    param(
        [Parameter(Mandatory)]$TestLD,
        [Parameter(Mandatory)]$NewDirectionName,
        [Parameter(Mandatory)]$exportCSV,
        [Parameter(Mandatory)]$ErrorTabCSV
    )

    #Define CSV Name
    $OldExportCSVFilename = $ExportCSV.split("\")[-1]
    $NewExportCSVFilename = $OldExportCSVFilename.replace($OldExportCSVFilename, "(Set)$OldExportCSVFilename")

    $OldErrorTabCSVFileName = $ErrorTabCSV.split("\")[-1]
    $NewErrorTabCSVFileName = $OldErrorTabCSVFileName.replace($OldErrorTabCSVFileName, "(Set)$OldErrorTabCSVFileName")

    $ExportCSV = $ExportCSV.replace($OldExportCSVFilename, $NewExportCSVFilename)
    $ErrorTabCSV = $ErrorTabCSV.replace($OldErrorTabCSVFileName,$NewErrorTabCSVFileName)

    $TabProperties = @{
        OldDisplayName = $null
        NewDisplayName = $null
        OldName =$null
        NewName = $null
        ObjectGuid = $null
        OldPrincipalAddress = $null
        NewPrincipalAddress = $null
        OldSecondaryAddress = $null
        NewSecondaryAddress = $null
        CanonicalName = $null
        EmployeeType = $null
        OldDepartment = $null
        NewDepartment = $null
        OldExtensionAttribute4 = $null
        NewExtensionAttribute4 = $null
        OldSamAccountName = $null
        NewSamAccountName = $null
        "NewPrincipalAddress set" = $null
        "NewSecondaryAddress set" = $null
        "NewDisplayName set" = $null
        "NewDepartment set" = $null
        "NewExtensionAttribute4 set" = $null
        "NewSamAccountName set" = $null
        "NewName set" = $null
        "MailRule Set"=$null
    }

    ##########################   DEBUT  ########################## 
    #Security in case the technician is not using the good CSV
    if ($TestLD.NewDepartment -notlike $NewDirectionName){
        Write-Host "/!\ [ERROR] MAKE SURE YOU ARE SETTING THE GOOD DIRECTION /!\"
        break
    }
    $Tabline = New-Object psobject -Property $Tabproperties

    #Setting SamAccountName
    $Tabline = Set-SpecialAttribute -LD $TestLD -Attribute "sAMAccountName" -OldOrNew New -Tabline $Tabline

    #Setting Name
    $Tabline = Rename-LD-Name -LD $TestLD -ErrorTabCSV $ErrorTabCSV -Tabline $Tabline -OldOrNew New
        
    #Set Principal & Secondary Address 
    $Tabline = Set-Addresses -LD $TestLD -Tabline $Tabline -ErrorTabCSV $ErrorTabCSV -OldOrNew New -dc $dc

    #Get AD Object to modify
    $ADObj = Get-ADGroup -Server <domain> -Identity $TestLD.ObjectGUID -Properties * 


    #Setting ExtensionAttribute4
    $ADObj = Set-AdObjAttribute -Adobj $ADObj -Obj $TestLD -Attribute "ExtensionAttribute4" -Value $TestLD."NewExtensionAttribute4" -OldOrNew New


    #Setting Department
    $ADObj = Set-AdObjAttribute -Adobj $ADObj -Obj $TestLD -Attribute "Department" -Value $TestLD."NewDepartment" -OldOrNew New


    #Setting DisplayName 
    $ADObj = Set-AdObjAttribute -Adobj $ADObj -Obj $TestLD -Attribute "DisplayName" -Value $TestLD."NewDisplayName" -OldOrNew New

    #Setting DisplayNamePrintable 
    $ADObj = Set-AdObjAttribute -Adobj $ADObj -Obj $TestLD -Attribute "DisplayNamePrintable" -Value $null -OldOrNew New


    #region Apply settings to ADObject
    try{
        Set-ADGroup -Server <domain> -Instance $ADObj -ErrorAction Stop -WarningAction SilentlyContinue
    }
    catch{
        Set-ErrorTabline -LD $TestLD -ErrorType "ERROR" -ErrorDescription "Apply NewDisplayName,NewDepartment,NewExtensionAttribute4 to ADObject" -ErrorMessage $_.exception.message | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
        continue
    }
    #endregion Apply Settings to ADObject


    #DisplayName in tab
    $Tabline.NewDisplayName = $TestLD.NewDisplayName
    $Tabline.OldDisplayName = $TestLD.OldDisplayName
    $Tabline.'NewDisplayName set' = "Yes"

    #ExtensionAttribute4 in tab
    $Tabline.NewExtensionAttribute4 = $TestLD.NewExtensionAttribute4
    $Tabline.OldExtensionAttribute4 = $TestLD.OldExtensionAttribute4
    $Tabline."NewExtensionAttribute4 Set" = "Yes"

    #Department in tab
    $Tabline.NewDepartment = $TestLD.NewDepartment
    $tabline.OldDepartment = $TestLD.OldDepartment
    $Tabline."NewDepartment Set" = "Yes"

    #Others in tab
    $Tabline.ObjectGuid = $TestLD.ObjectGuid
    $Tabline.CanonicalName = $TestLD.CanonicalName



    #Write in console
    write-host "[SUCCESS] DisplayName : $($TestLD.NewDisplayName) | ObjectGUID : $($TestLD.ObjectGUID)" -ForegroundColor green



    #End of set, export CSV

    $Tabline | Select-Object OldDisplayName, NewDisplayName, "NewDisplayName set", OldPrincipalAddress, NewPrincipalAddress, "NewPrincipalAddress set", OldSecondaryAddress, NewSecondaryAddress, "NewSecondaryAddress set", OldSamAccountName, NewSamAccountName, "NewSamAccountName set", OldName, NewName, "NewName set", OldDepartment, NewDepartment, "NewDepartment set", OldExtensionAttribute4, NewExtensionAttribute4, "NewExtensionAttribute4 set", "MailRule Set", ObjectGuid, CanonicalName | Export-Csv -Path $exportCSV -Encoding UTF8 -Delimiter ";" -NoTypeInformation -Append
    Return $Tabline | Select-Object OldDisplayName, NewDisplayName, "NewDisplayName set", OldPrincipalAddress, NewPrincipalAddress, "NewPrincipalAddress set", OldSecondaryAddress, NewSecondaryAddress, "NewSecondaryAddress set", OldSamAccountName, NewSamAccountName, "NewSamAccountName set", OldName, NewName, "NewName set", OldDepartment, NewDepartment, "NewDepartment set", OldExtensionAttribute4, NewExtensionAttribute4, "NewExtensionAttribute4 set", "MailRule Set", ObjectGuid, CanonicalName
}

Function Restore-LDs
{
    param(
        [Parameter(Mandatory)]$SetLD,
        [Parameter(Mandatory)]$OldDirectionName,
        [Parameter(Mandatory)]$exportCSV,
        [Parameter(Mandatory)]$ErrorTabCSV
    )
    #Define CSV Name
    $OldExportCSVFilename = $ExportCSV.split("\")[-1]
    $NewExportCSVFilename = $OldExportCSVFilename.replace($OldExportCSVFilename, "(Restore)$OldExportCSVFilename")

    $OldErrorTabCSVFileName = $ErrorTabCSV.split("\")[-1]
    $NewErrorTabCSVFileName = $OldErrorTabCSVFileName.replace($OldErrorTabCSVFileName, "(Restore)$OldErrorTabCSVFileName")

    $ExportCSV = $ExportCSV.replace($OldExportCSVFilename, $NewExportCSVFilename)
    $ErrorTabCSV = $ErrorTabCSV.replace($OldErrorTabCSVFileName,$NewErrorTabCSVFileName)

    $TabProperties = @{
        OldDisplayName = $null
        NewDisplayName = $null
        OldName =$null
        NewName = $null
        ObjectGuid = $null
        OldPrincipalAddress = $null
        NewPrincipalAddress = $null
        OldSecondaryAddress = $null
        NewSecondaryAddress = $null
        CanonicalName = $null
        EmployeeType = $null
        OldDepartment = $null
        NewDepartment = $null
        OldExtensionAttribute4 = $null
        NewExtensionAttribute4 = $null
        OldSamAccountName = $null
        NewSamAccountName = $null
        "OldPrincipalAddress set" = $null
        "OldSecondaryAddress set" = $null
        "OldDisplayName set" = $null
        "OldDepartment set" = $null
        "OldExtensionAttribute4 set" = $null
        "OldSamAccountName set" = $null
        "OldName set" = $null
        "MailRule Set" = $null
    }

    $Tabline = New-Object psobject -Property $Tabproperties

    #Setting SamAccountName
    $Tabline = Set-SpecialAttribute -LD $SetLD -Attribute "sAMAccountName" -OldOrNew Old -Tabline $Tabline

    #Setting Name
    $Tabline = Rename-LD-Name -LD $SetLD -ErrorTabCSV $ErrorTabCSV -Tabline $Tabline -OldOrNew Old
        
    #Set Principal & Secondary Address 
    $Tabline = Set-Addresses -LD $SetLD -Tabline $Tabline -ErrorTabCSV $ErrorTabCSV -OldOrNew Old -dc $dc


    $ADObj = Get-ADGroup -Server <domain> -Identity $SetLD.ObjectGUID -Properties * 

    #Setting ExtensionAttribute4
    $ADObj = Set-AdObjAttribute -Adobj $ADObj -Obj $SetLD -Attribute "ExtensionAttribute4" -Value $SetLD."OldExtensionAttribute4" -OldOrNew Old


    #Setting Department
    $ADObj = Set-AdObjAttribute -Adobj $ADObj -Obj $SetLD -Attribute "Department" -Value $SetLD."OldDepartment" -OldOrNew Old


    #Setting DisplayName 
    $ADObj = Set-AdObjAttribute -Adobj $ADObj -Obj $SetLD -Attribute "DisplayName" -Value $SetLD."OldDisplayName" -OldOrNew Old


    #Setting DisplayNamePrintable 
    $ADObj = Set-AdObjAttribute -Adobj $ADObj -Obj $SetLD -Attribute "DisplayNamePrintable" -Value $null -OldOrNew Old

    #region Apply settings to ADObject
    try{
        Set-ADGroup -Server <domain> -Instance $ADObj -ErrorAction Stop -WarningAction SilentlyContinue
    }
    catch{
        Set-ErrorTabline -LD $SetLD -ErrorType "ERROR" -ErrorDescription "Apply NewDisplayName,NewDepartment,NewExtensionAttribute4 to ADObject" -ErrorMessage $_.exception.message | Export-Csv $ErrorTabCSV -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
        continue
    }
    #endregion Apply Settings to ADObject
    
    
    
    #DisplayName in tab
    $Tabline.NewDisplayName = $SetLD.NewDisplayName
    $Tabline.OldDisplayName = $SetLD.OldDisplayName
    $Tabline.'OldDisplayName set' = "Yes"

    #ExtensionAttribute4 in tab
    $Tabline.NewExtensionAttribute4 = $SetLD.NewExtensionAttribute4
    $Tabline.OldExtensionAttribute4 = $SetLD.OldExtensionAttribute4
    $Tabline."OldExtensionAttribute4 Set" = "Yes"

    #Department in tab
    $Tabline.NewDepartment = $SetLD.NewDepartment
    $tabline.OldDepartment = $SetLD.OldDepartment
    $Tabline."OldDepartment Set" = "Yes"

    #Others in tab
    $Tabline.ObjectGuid = $SetLD.ObjectGuid
    $Tabline.CanonicalName = $SetLD.CanonicalName



    #Write in console
    write-host "[SUCCESS] DisplayName : $($SetLD.OldDisplayName) | ObjectGUID : $($SetLD.ObjectGUID)" -ForegroundColor green

    
    $Tabline | Select-Object OldDisplayName, NewDisplayName, "OldDisplayName set", OldPrincipalAddress, NewPrincipalAddress, "OldPrincipalAddress set", OldSecondaryAddress, NewSecondaryAddress, "OldSecondaryAddress set", OldSamAccountName, NewSamAccountName, "OldSamAccountName set", OldName, NewName, "OldName set", OldDepartment, NewDepartment, "OldDepartment set", OldExtensionAttribute4, NewExtensionAttribute4, "OldExtensionAttribute4 set", "MailRule Set", ObjectGuid, CanonicalName, EmployeeType | Export-Csv -Path $exportCSV -Encoding UTF8 -Delimiter ";" -NoTypeInformation -Append
    Return $Tabline | Select-Object OldDisplayName, NewDisplayName, "OldDisplayName set", OldPrincipalAddress, NewPrincipalAddress, "OldPrincipalAddress set", OldSecondaryAddress, NewSecondaryAddress, "OldSecondaryAddress set", OldSamAccountName, NewSamAccountName, "OldSamAccountName set", OldName, NewName, "OldName set", OldDepartment, NewDepartment, "OldDepartment set", OldExtensionAttribute4, NewExtensionAttribute4, "OldExtensionAttribute4 set", "MailRule Set", ObjectGuid, CanonicalName, EmployeeType
}









################################################################# Dï¿½but #########################################################################
#Test if the technician who's launching this script is part of the good groups

$ADGroups_T1_T0 = <Groups>
$UserRights = $ADGroups_T1_T0 | foreach-object {Get-ADGroup -server <domain> -Identity $_ -Properties member | Where-Object member -like "CN=$env:USERNAME*"}
if (!$UserRights){
    write-host "Le compte $env:USERNAME n'a pas les droits nécéssaire pour effectuer les actions de ce script" -ForegroundColor Red
    break
}


#Set time in file
$filetime = Get-Date -Format "dd-MM-yyyy_HH-mm-ss"
$exportCSV = $exportCSV.replace(".csv", "_$filetime.csv")
$ErrorTabCSV = $ErrorTabCSV.replace(".csv", "_$filetime.csv")


###### TEST NEW CONFIGURATION ######
if ($Test){
    $TestTab = @()
    $LDs = Get-LDs -OldDirectionName $OldDirectionName -DirectionOU $DirectionOU
    $Counter = 0
    $EndCounter = $LDs.Count
    foreach ($LD in $LDs) {
        if ($EndCounter -gt 1){
            Write-Progress -Activity "Simulate LDs ($Counter/$endCounter)" -PercentComplete $($Counter++ / $endCounter * 100) -Id 1
        }
        $TestTab += Test-LD-Modifications -LD $LD -OldDirectionName $OldDirectionName -NewDirectionName $NewDirectionName -ExportCSV $ExportCSV -ErrorTabCSV $ErrorTabCSV

    }
    $TestTab | Out-GridView
    Return $TestTab
}


###### SET NEW CONFIGURATION ######
if ($Set){
    $SetTab = @()
    Write-Host "/!\ --- DON'T FORGET TO ENSURE THAT THE RESULTS ARE GOOD --- /!\" -ForegroundColor Yellow
    $Continue = Read-Host -Prompt "/!\ After you have seen the test table do you want to apply the modifications ? Yes - No "
    if ($Continue -eq "Yes"){

        $TestTab = Import-Csv $TestFile -Delimiter ";" -Encoding UTF8
        if(($TestTab.OldName | Select-String "�") -or ($TestTab.OldDisplayName | Select-String "�") -or ($TestTab.OldSamAccountName | Select-String "�")){

            Return Write-Host "[ERROR] Le Tableau CSV contient des caractères illisibles (ex: �), celà peut être dû à l'enregistrement via Excel" -ForegroundColor Red

        }

        $Counter = 0
        $EndCounter = $TestTab.Count
        foreach ($Testline in $TestTab){
            if ($EndCounter -gt 1){
                Write-Progress -Activity "Configure LDs ($Counter/$endCounter)" -PercentComplete $($Counter++ / $endCounter * 100) -Id 1
            }
            $SetTab += Set-LDs-Modifications -TestLD $Testline -NewDirectionName $NewDirectionName -exportCSV $ExportCSV -ErrorTabCSV $ErrorTabCSV
        }
    
    } 
    $SetTab | Out-GridView 
    Return $SetTab
}


###### RESTORE OLD CONFIGURATION ######
if ($Restore){
    $SetFileToRestore = Import-Csv $SetFileToRestore -Delimiter ";" -Encoding UTF8
    $RestoreTab = @()
    $Counter = 0
    $EndCounter = $SetFileToRestore.Count

    Foreach ($LD in $SetFileToRestore){
        if ($EndCounter -gt 1){
            Write-Progress -Activity "Restore LDs ($Counter/$endCounter)" -PercentComplete $($Counter++ / $endCounter * 100) -Id 1
        }
        $RestoreTab += Restore-LDs -SetLD $LD -OldDirectionName $OldDirectionName -exportCSV $ExportCSV -ErrorTabCSV $ErrorTabCSV
    }
    $RestoreTab | Out-GridView
    Return $RestoreTab
}
