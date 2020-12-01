Write-Host 'Install PowerApps Powershell Modules'
Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -AllowClobber -Scope CurrentUser -Force
Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber  -Scope CurrentUser -Force
Install-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell -AllowClobber -Scope CurrentUser -Force
Install-Module -Name Microsoft.Xrm.Data.PowerShell -AllowClobber -Scope CurrentUser -Force

#EXAMPLES
#ShareSolutionWithUser -ServerUrl "https://5mpp-pcdev2.crm.dynamics.com" -SolutionName "PowerCATs" -UserName "user@domain.com"
#ConnectToPowerAppEnvironment -Username $user -Password $pass -ServerUrl "https://5mpp-pcdev2.crm.dynamics.com"
#Set-PowerAppEnvironmentVariable -SchemaName "fmpp_CatNotifyEmail" -Value "newenvvalue"
#Write-Host $(Get-PowerAppEnvironmentVariable -SchemaName "fmpp_CatNotifyEmail")

#Assumes these PS environment variables are set
#$env:SERVICEACCOUNTPASSWORD
#$env:SERVICEACCOUNTNAME

#Globals
#Microsoft.Xrm.Data.PowerShell sets a global CDS connection $conn, which is used implicitly throughout
#Calling ConnectToPowerAppEnvironment will set $conn
#$pass = ConvertTo-SecureString -string $env:SERVICEACCOUNTPASSWORD -AsPlainText -Force
#$user = $env:SERVICEACCOUNTNAME
   
#ShareSolutionWithUser also connects to CDS and sets $conn
Function ShareSolutionWithUser
{ 
    Param ([String]$ServerUrl, [String]$SolutionName, [String]$UserName)
    
    ConnectToPowerAppEnvironment -Username $user -Password $pass -ServerUrl $ServerUrl

    #User Setup
    Add-PowerAppsAccount -Username $user -Password $pass -Endpoint prod
    $userGraph = Get-UsersOrGroupsFromGraph -SearchString $UserName
    $userXrmResults = Get-CrmRecords -EntityLogicalName "systemuser" -FilterAttribute domainname -FilterOperator eq -FilterValue $UserName -Fields fullname 
    
    if ($userXrmResults.Count -ne 0 -and $userGraph -ne $null)
    {
        $userXrmRecord = $userXrmResults.CrmRecords[0]
    }
    else
    {
        Write-Error "No user matching $UserName found."
        Throw
    }
  
    #Retrieve Solution ID
    $solutionXrmResults = (Get-CrmRecords -EntityLogicalName solution -FilterAttribute uniquename -FilterOperator eq -FilterValue $SolutionName)
    if($solutionXrmResults.Count -eq 1)
    {
        $solutionId = $solutionXrmResults.CrmRecords[0].solutionid
    } 
    else
    {
        
        Write-Error "No solution matching $SolutionName found."
        Throw
    }


    ShareSolutionComponents $solutionId "environmentvariabledefinition" "schemaname" $userXrmRecord

    ShareSolutionComponents $solutionId "workflow" "name" $userXrmRecord

    foreach($canvasApp in Get-SolutionComponents $solutionId "canvasapp" "displayname")
    {
        Set-PowerAppRoleAssignment -PrincipalType User -PrincipalObjectId $userGraph.ObjectId -RoleName CanEdit -AppName $canvasApp.canvasappid -EnvironmentName $conn.EnvironmentId
        Write-Host "Shared Canvas App $($canvasapp.displayname)"
    }

    foreach($role in Get-SolutionComponents $solutionid "role" "name")
    {
        try{
            Add-CrmRecordAssociation -EntityLogicalName1 $role.logicalname -Id1 $role.roleid -EntityLogicalName2 $userXrmRecord.LogicalName -Id2 $userXrmRecord.EntityReference.Id -RelationshipName "systemuserroles_association" 
            Write-Host "Added role $($role.name) to $($userXrmRecord.fullname)"
        }
        catch [System.Management.Automation.RuntimeException]
        {
            Write-Host "User $($userXrmRecord.fullname) already has security role $($role.name)"
        }
        catch
        {
            Throw $Error[0].Exception
        }
    }

}

#Helper functions to iterate records by solution
#Use global variable $conn from Microsoft.Xrm.Data.PowerShell

Function Get-SolutionComponents {
    Param ([String]$SolutionID, [String]$EntityLogicalName, [String]$NameField)
    Return (Get-CrmRecords -EntityLogicalName $EntityLogicalName -FilterAttribute solutionid -FilterOperator eq -FilterValue $SolutionID -Fields $NameField).CrmRecords
}

Function ShareSolutionComponents{
    Param ([String]$SolutionID, [String]$EntityLogicalName, [String]$NameField, [PSObject]$UserObj)
    foreach($component in Get-SolutionComponents $SolutionID $EntityLogicalName $NameField)
    {
        Grant-CrmRecordAccess -CrmRecord $component -Principal $UserObj.EntityReference -AccessMask ReadAccess,WriteAccess
        Write-Host Shared $component.logicalname $component.$NameField 'with user' $UserObj.fullname
    }
}

Function ConnectToPowerAppEnvironment {
    Param(
        [String]$Username,
        [Security.SecureString]$Password,
        [String]$ServerUrl
        )

    try
    {
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $Password
        Connect-CrmOnline -Credential $cred -ServerUrl $ServerUrl 
    }
    catch
    {
        Write-Error "Unable to connect to $ServerUrl."
        Throw
    }

    if (!$conn.IsReady)
    {
        Write-Error "No connection to $ServerUrl - $($conn.LastCrmError)"
        Throw $conn.LastCrmException
    }
    return $conn
}

Function Set-PowerAppEnvironmentVariable{
#Uses global connection variable $conn
    Param(
    [String]$SchemaName, 
    [String]$Value
    )

    $envVarDefResults = Get-CrmRecords -EntityLogicalName "environmentvariabledefinition" -FilterAttribute "schemaname" -FilterOperator eq -FilterValue $SchemaName

    if($envVarDefResults.Count -eq 1)
    {
        $EnvVarValResults = Get-CrmRecords -EntityLogicalName "environmentvariablevalue" -FilterAttribute "environmentvariabledefinitionid" -FilterOperator eq -FilterValue $envvardefresults.CrmRecords[0].environmentvariabledefinitionid -Fields "value"

        if($EnvVarValResults.Count -ne 0)
        {
            $EnvVarRecord = $EnvVarValResults.CrmRecords[0]
            $EnvVarRecord.value = $Value
            Set-CrmRecord -CrmRecord $EnvVarRecord
        }
        else
        {
            New-CrmRecord -EntityLogicalName "environmentvariablevalue" -Fields @{"value"=$Value; "environmentvariabledefinitionid"=$envvardefresults.CrmRecords[0].EntityReference}

        }
    
    }
    else
    {
        Write-Error "Found $($EnvVarValResults.Count) enviornment variable definitions with schema name $SchemaName"
    }
}

Function Get-PowerAppEnvironmentVariable{
#Uses global connection variable $conn
    Param(
    [String]$SchemaName
    )

    $envVarDefResults = Get-CrmRecords -EntityLogicalName "environmentvariabledefinition" -FilterAttribute "schemaname" -FilterOperator eq -FilterValue $SchemaName

    if($envVarDefResults.Count -eq 1)
    {
        $EnvVarValResults = Get-CrmRecords -EntityLogicalName "environmentvariablevalue" -FilterAttribute "environmentvariabledefinitionid" -FilterOperator eq -FilterValue $envvardefresults.CrmRecords[0].environmentvariabledefinitionid -Fields "value"
        $EnvVarRecord = $EnvVarValResults.CrmRecords[0]
        return $EnvVarRecord.value
    }
    else
    {
        Write-Error "Found $($EnvVarValResults.Count) enviornment variable definitions with schema name $SchemaName"
    }

}


