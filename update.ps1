
#Initialize default properties
$c = $configuration | ConvertFrom-Json
$p = $person | ConvertFrom-Json;
$pp = $previousPerson | ConvertFrom-Json
$pd = $personDifferences | ConvertFrom-Json
$m = $manager | ConvertFrom-Json;
$aRef = $accountReference | ConvertFrom-Json;
$mRef = $managerAccountReference | ConvertFrom-Json;
$success = $False;
$auditMessage = $p.DisplayName;

#Retrieve current account data for properties to be updated
$previousAccount = [PSCustomObject]@{
    email = $pp.Accounts.ActiveDirectory.mail;
}

if($pp.Accounts.ActiveDirectory.mail -eq $p.Accounts.ActiveDirectory.mail)
{
    #No Changes Required
    Write-Verbose -Verbose "No Changes needed."
    $success = $False;
    $auditMessage = $p.DisplayName;
}
else
{
    #Authorization
        function get_oauth_access_token {
        [cmdletbinding()]
        Param (
        [string]$BaseURI,
        [string]$ClientKey,
        [string]$ClientSecret
        )
            Process
            {
                $pair = $ClientKey + ":" + $ClientSecret;
                $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair);
                $bear_token = [System.Convert]::ToBase64String($bytes);
                $auth_headers = @{ Authorization = "Basic " + $bear_token };
                
                $uri =  "$($BaseURI)/oauth/token?grant_type=client_credentials";
                $result = Invoke-RestMethod -Method GET -Headers $auth_headers -Uri $uri -UseBasicParsing;
                @($result);
            }
        }
    
    #Update Record
    try {         
        $AccessToken = (get_oauth_access_token `
                    -BaseURI $c.BaseURI `
                    -ClientKey $c.ClientKey `
                    -ClientSecret $c.ClientSecret).access_token

        $headers = @{ Authorization = "Bearer $($AccessToken)"; "Content-Type"= "application/json" };

        #Get Existing Email Address
            
            #Email Type
            #See /Generic/1/Demographics/EmailType/1/100 for Types
            $EmailTypeID = 1 

            #Email Rank
            # 1 = Primary
            $EmailRank = 1

            #Search for record encase old one was deleted
            $uri = "$($c.BaseURI)/Generic/1/Demographics/NameEmail/1/100?searchFields[0]=NameEmailID&searchFields[1]=EmailAddress&searchFields[2]=EmailTypeID&searchFields[3]=NameID&searchFields[4]=Rank"
            $body = @{
                        "SearchCondition" = @{
                            "SearchConditionGroup" = @{
                            "ConditionGroupType"= "And";
                            "Conditions"= @(
                                    @{
                                        #Filter Results by NameID
                                        "LongSearchCondition"= @{
                                            "ConditionType"= "Equal";
                                            "FieldName"= "NameID";
                                            "Value"= $p.ExternalID;
                                        }
                                    },
                                    @{
                                        #Filter by Email Type
                                        "LongSearchCondition"= @{
                                            "ConditionType"= "Equal";
                                            "FieldName"= "EmailTypeID";
                                            "Value"= $EmailTypeID;
                                        }
                                    },
                                    @{
                                        #Filter by Email Rank
                                        "LongSearchCondition"= @{
                                            "ConditionType"= "Equal";
                                            "FieldName"= "Rank";
                                            "Value"= $EmailRank;
                                        }
                                    }
                                )
                            }
                        }
            }

            $currentRecords = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body ($body | ConvertTo-Json -Depth 50)

            #Check if Records Exist
            if($currentRecords.Objects.Count -gt 0)
            {
                Write-Verbose -Verbose "Found Existing Email Record";
                #Use existing record, update
                $account = [PSCustomObject]@{
                    DataObject = $currentRecords.Objects[0]
                }
            }
            else
            {
                #No existing records, create new record
                Write-Verbose -Verbose "No Existing Email Record Found"
                $account = [PSCustomObject]@{
                    DataObject = @{
                        EmailAddress = "";
                        EmailTypeID = $EmailTypeID;
                        NameID = $p.ExternalID;
                        Rank = $EmailRank;
                    }
                }
            }

        #Change mapping here
        $account.DataObject.EmailAddress = $p.Accounts.ActiveDirectory.mail;

        if(-Not($dryRun -eq $True)) {
            if($account.DataObject.NameEmailID -eq $null)
            {
                #Create
                Write-Verbose -Verbose "Creating new email record";
                $uri = "$($c.BaseURI)/Generic/1/Demographics/NameEmail?searchFields[0]=NameEmailID&searchFields[1]=EmailAddress&searchFields[2]=EmailTypeID&searchFields[3]=NameID&searchFields[4]=Rank"
                $result = Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -Body ($account | ConvertTo-Json -Depth 20) -ContentType "application/json" -UseBasicParsing;
                $aRef = $result.NameEmailID;
            }
            else
            {
                #Update
                Write-Verbose -Verbose "Updating existing email record";
                $uri = "$($c.BaseURI)/Generic/1/Demographics/NameEmail/$($account.DataObject.NameEmailID)?searchFields[0]=NameEmailID&searchFields[1]=EmailAddress&searchFields[2]=EmailTypeID&searchFields[3]=NameID&searchFields[4]=Rank"
                $result = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body ($account | ConvertTo-Json -Depth 20) -ContentType "application/json" -UseBasicParsing;
                $aRef = $result.NameEmailID;
            }
        }
        $success = $True;
        $auditMessage = " successfully"; 
    }catch{
        $auditMessage = " : General error $($_)";
        Write-Error -Verbose $_; 
    }
}

#build up result
$result = [PSCustomObject]@{
	Success= $success;
	AuditDetails=$auditMessage;
    Account = $account;
    PreviousAccount = $previousAccount;
};

Write-Output $result | ConvertTo-Json -Depth 10;
