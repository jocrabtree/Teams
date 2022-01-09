# Teams_NoOwners
# Author: Josh Crabtree, Exeter Finance, 8 Jan 2022

#region variables
$ExactTime = Get-Date -Format "MM-dd-yyyy HHmm tt"
$Today = (Get-Date).ToString('MM-dd-yyyy')
$SMTP = "<YOUR SMTP ADDRESS HERE>"
$global:ErrorActionPreference = 'Stop'

# Credential Setup
$UN = "<YOUR ADMIN USERNAME HERE>"
$Key = Get-Content "<AES KEY FILE HERE>"
$PW = Get-Content "<PASSWORD IN A FILE HERE>" | ConvertTo-SecureString -Key $Key
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $UN, $PW

#Variables used to send report
$To = "<YOUR EMAIL HERE>"
$Cc = "<CC ADDRESS HERE>"
$Bcc = "<YOUR BCC HERE>"
$logfile = "C:\TeamsWithNoOwners-$($Today).log"
$TeamsNoOwnersCSV = "C:\TeamsWithNoOwners-$Today.csv"

#endregion variables

#region functions

#Function used for writing to the log file
function Log-It {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $True,
            Position = 0,
            ValueFromPipeline=$True
        )]
        [String]$Message,
        [ValidateSet(
            "General","Process","Success","Failure","Warning","Notification","LogOnly","ScreenOnly"
        )]
        [String]$Status = "General"
    )
    Switch($Status){
        "General"{
            $Color="Cyan"
            $Type="[INFORMA] "
        }
        "Process"{
            $Color="White"
            $Type="[PROCESS] "
        }
        "Failure"{
            $Color="Red"
            $Type="[FAILURE] "
        }
        "Success"{
            $Color="Green"
            $Type="[SUCCESS] "
        }
        "Warning"{
            $Color="Yellow"
            $Type="[WARNING] "
        }
        "Notification"{
            $Color="Gray"
            $Type="[NOTICES] "
        }
        "ScreenOnly"{
            $Color="Magenta"
            $Type="[INFORMA] "
        }
        "LogOnly"{
            $Color=$Null
            $Type="[INFORMA] "
        }

    }
    if($Color){Write-Host -ForegroundColor $Color $Type$Message}
    if($Color -ne "Magenta"){"$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Type$Message" | Out-File $logfile -Append}
}

#Function used to connect to Exchange Online so I can pull in M365 Group information.
function Connect-EXCOnline {
    $FN = "Connect-EXCOnline"
    try{
        $EOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection 
        $Import = Import-PSSession $EOLSession -AllowClobber -WarningVariable ignore -InformationAction Ignore -DisableNameChecking
        "$FN | Connected to Exchange Online" | Log-It -Status Success
    }
    catch{
        "$FN | Failed to connect to Exchange Online! | $($error[0].exception.message)" | Log-It -Status Failure
    }
}

#Function to find all team members of teams w/o owners
function Get-NoOwnerTeamMembers{
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true
        )]
        $Team
    )
    begin{
        $Results = @()
        $FN = "Get-NoOwnerTeamMembers"
        "$FN | Finding members of all Teams with no owners" | Log-it -Status Notification
    }
    process{
        try{
            $GroupMembers = Get-UnifiedGroupLinks -Identity $_.DisplayName -LinkType Members
            if($GroupMembers){
            "Found Members for Team: $($_.DisplayName)"|Log-it -Status Success
                $Results += [pscustomobject]@{
                    TeamName = $_.DisplayName;
                    Email = $_.PrimarySmtpAddress;
                    Members = $GroupMembers.Name -join ',';
                    CreationDate = $_.WhenCreated;
                }
            }
            else{
                "Could not find members for Team: $($_.DisplayName)"|Log-it -Status General
                $Results += [pscustomobject] @{
                    TeamName = $_.DisplayName;
                    Email = $_.PrimarySmtpAddress;
                    Members = "No Group Members Found";
                    CreationDate = $_.WhenCreated;
                }
            }
        }
        catch{
            "There was an issue finding members for this Team: $($_.displayname) | $($error[0].exception.message)" | Log-It -Status Failure 
        }
    }
    end{
        if(!$Results){"No Teams were found without owners." | Log-It -Status Notification}
        else{
            Return @{Results=$Results;}
        }
    }
}

#Function to build a new report
function New-Report {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline=$true
        )]
        [ValidateNotNullOrEmpty()]
        $User,
        [Parameter(
            Mandatory = $true,
            Position = 1
        )]
        [ValidateNotNullOrEmpty()]
        $Topic,
        [Parameter(
            Mandatory = $False,
            Position = 2
        )]
        $TableName,
        [Parameter(
            Mandatory = $False,
            Position = 3
        )]
        $Notes,
        [Parameter(
            Mandatory = $False,
            Position = 4
        )]
        $Properties = @("Name","UserPrincipalName","SamAccountName","Mail","Title","Department","Manager","DistinguishedName"),
        [Parameter(
            Mandatory = $False,
            Position = 5
        )]
        $CustomProps,
        [Parameter(
            Mandatory = $False,
            Position = 6
        )]
        $Count
    )
    begin{
        $Table = @()
        $FN = "New-Report"
        Log-It "[OK] $FN | Generating Report for $Topic"
        $AllProps = $Properties 
        if($CustomProps){$AllProps += $CustomProps}
        $TableProps = $AllProps | Where {$_ -notlike "*HTML"}
        $HTMLProps = $AllProps | Where {$_ -notlike "*Table" -and $_ -ne "UserPrincipalName"}
        $HTML = @" 
            $(if($TableName){"<h3>$TableName</h3>"})
            $(if($Notes){"<p>$Notes</p>"})
            <br>
            $(if($Count){"<p><strong>Total:<u>$Count</u></strong></p>"})
            <table>
            <thead>
            <tr>
"@
        $HTMLProps.ForEach({
            $Header = $_
            if($Header -eq "DistinguishedName"){
                $Header = "OU"
            }
            elseif($Header -like "*HTML"){
                $Header = $Header.Replace("HTML","")
            }
            $HTML += "<th>$Header</th>"
        })
        $HTML += "</tr>`n"
        $HTML += @"
            </thead>
            <tbody>
"@
    }
    Process{
        #Generate Table
        $Table += $User | select $TableProps
        $HTML += "<tr>"
        ForEach($Header in $HTMLProps){
            if(($User.$Header)){
                # AD User Reporting
                if($User.Gettype().name -eq "ADUser"){
                    if($Header -eq "DistinguishedName"){
                        $OU = ($User.$Header.split(",") | Select-Object -Skip 1) -join ","
                        $HTML += "<td>$($OU)</td>"
                    }
                    elseif($Header -eq "Manager"){
                        $Manager = $($User.$Header).replace("CN=","").split(",")[0]
                        $HTML += "<td>$($Manager)</td>"
                    }
                    elseif($Header -eq "Created"){
                        $Created = $User.Created
                        $HTML += "<td>$($Created)</td>"
                    }
                    elseif($Header -like "Enabled"){
                        if($User.$Header){
                            $tde = '<td style="color:green">'
                            $Enabled = "TRUE"
                        }
                        else{
                            $tde = '<td style="color:red">'
                            $Enabled = "FALSE"
                        }
                        $HTML += "$tde$($Enabled)</td>"
                    }
                    else{
                        $HTML += "<td>$($User.$Header)</td>"
                    }
                }
                elseif($Header -eq "CustomAttribute12"){
                    $SAM = $User.$Header
                    $Manager = (Get-ADUser -f{SamAccountName -like $SAM}).Name
                    $HTML += "<td>$($Manager)</td>"
                }
                else{
                    $HTML += "<td>$($User.$Header)</td>"
                }
            }
            else{
                $HTML += '<td style="color:red">No Value</td>'
            }
        }
        $HTML += "</tr>`n"
    }
    end{
        $HTML += @"
            </tbody>
            </table>    
            <br>   
"@
        return @{Table=$Table;HTML=$HTML}
    }    
}

#Function to email the report
function Email-Report {
    [Cmdletbinding()]
    param(
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true            
        )]
        $HTMLBody,
        [Parameter(
            Mandatory = $true,
            Position = 1
        )]
        $MailTo,
        [Parameter(
            Mandatory = $false,
            Position = 2
        )]
        $Cc,
        [Parameter(
            Mandatory = $false,
            Position = 3
        )]
        $Bcc,
        [Parameter(
            Mandatory = $true,
            Position = 4
        )]
        $SMTP,
        [Parameter(
            Mandatory = $False,
            Position = 5
        )]
        $Attachments,
        [Parameter(
            Mandatory = $False,
            Position = 6
        )]
        $HTMLAttach,
        [Parameter(
            Mandatory = $True,
            Position = 7
        )]
        $Title,
        [Parameter(
            Mandatory = $False,
            Position = 8
        )]
        [ValidateSet(
            "Reporting", "Training", "Deprovisioning", "Decommissioning"
        )]
        $Type,
           [Parameter(
            Mandatory = $False,
            Position = 9
        )]
        $Notes
    )
    begin{
        $FN = "Email-Report"
        "$FN | Emailing $Title Report" | Log-It -Status General
        switch($Type){
            "Reporting"{
                $SendAs = "<YOUR EMAIL HERE>"
                $Header='<img src="<YOUR IMAGE HERE>" width=2000 border="0">'
            };
            "Training"{
                $SendAs = "<YOUR EMAIL HERE>"
                $Header='<img src="<YOUR IMAGE HERE>" width=2000 border="0">'
            };
            "Deprovisioning"{
                $SendAs = "<YOUR EMAIL HERE>"
                $Header='<img src="<YOUR IMAGE HERE>" width=2000 border="0">'
            };
            "Decommissioning"{
                $SendAs = "<YOUR EMAIL HERE>"
                $Header='<img src="<YOUR IMAGE HERE>" width=2000 border="0">'
            };
            default{
                $SendAs = "<YOUR EMAIL HERE>"
                $Header='<img src="<YOUR IMAGE HERE>" width=2000 border="0">'
            };
        }
    }
    process{
        # Create HTML Report
        $HTML = @"
            <html>
                <head>
                    <style>
                        table{
                            font-size: 1em;
                            text-align: center;
                            border: 1px solid black;
                            margin: auto;
                            width: 90%;
                        }
                        th{
                            background-color: #4CD533;
                            color: white;
                            border: 1px #4CD533;
                        }
                        td{
                            border: 1px solid black;
                            border-collapse: collapse;
                        }
                        body{
                            box-sizing: border-box;                                                    
                            text-align: center;
                            font-family: Arial;

                        }
                        img{
                            display: block;
                            margin-left: auto;
                            margin-right: auto;
                            width: 100%;
                        }
                    </style>
                </head>
                <body>
	                <h1>$Header</h1>
                    <br>
                    <h2>
                        $Title
                    </h2>
                    <br>
                    $(if($Notes){$Notes})
                    $HTMLBody
                    <br>
                    <p><i>This automated report ran at $((Get-Date).ToString())</i></p>
                    <h1><img src="<YOUR IMAGE HERE>" width=2000 border="0"></h1>
                </body>    
            </html>
"@

        # Create HTML File to Attach
        if($HTMLAttach){
            $HTMLFile = "$($HTMLFilePath)$($Today)$($Title).HTML"
            $HTML | Out-File $HTMLFile
            if($Attachments){$Attachments += $HTMLFile}
            else{$Attachments = $HTMLFile}
        }

        # Email report with attached CSV file
        try{
            if($Attachments){
                if($Cc -and $Bcc){Send-MailMessage -From $SendAs -To $MailTo -Cc $Cc -Bcc $Bcc -Subject "$Title - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer $SMTP -Attachments $Attachments -ErrorAction Stop}
                elseif($Cc){Send-MailMessage -From $SendAs -To $MailTo -Cc $Cc -Subject "$Title - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer $SMTP -Attachments $Attachments -ErrorAction Stop}
                elseif($Bcc){Send-MailMessage -From $SendAs -To $MailTo -Bcc $Bcc -Subject "$Title - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer $SMTP -Attachments $Attachments -ErrorAction Stop}
                else{Send-MailMessage -From $SendAs -To $MailTo -Subject "$Title - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer $SMTP -Attachments $Attachments -ErrorAction Stop}
            }
            else{
                if($Cc -and $Bcc){Send-MailMessage -From $SendAs -To $MailTo -Cc $Cc -Bcc $Bcc -Subject "$Title - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer $SMTP -ErrorAction Stop}
                elseif($Cc){Send-MailMessage -From $SendAs -To $MailTo -Cc $Cc -Subject "$Title - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer $SMTP -ErrorAction Stop}
                elseif($Bcc){Send-MailMessage -From $SendAs -To $MailTo -Bcc $Bcc -Subject "$Title - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer $SMTP -ErrorAction Stop}
                else{Send-MailMessage -From $SendAs -To $MailTo -Subject "$Title - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer $SMTP -ErrorAction Stop}
            }
            "$FN | Successfully sent the $Title Report" | Log-It -Status Success
        }
        catch{
            "$FN | Failed to send the $Title Report | $($error[0].exception.message)." | Log-It -Status Failure
        }
    }
}

#endregion functions

#region process
"Begin Script" | Log-it -Status General

#Run function to Connect to Exchange Online
Connect-EXCOnline

#Array to store all Teams w/o owners
$NoOwners = @()

"Checking for Teams without owners" | Log-it -Status General

#Use the Get-UnifiedGroup command in Exchange Online w/ the ResourceProvisionOptions param set to Team to find all Teams w/o owners
$NoOwners += Get-UnifiedGroup | Where-Object {(-Not $_.ManagedBy) -and ($_.ResourceProvisioningOptions -eq "Team")}
"Found $($NoOwners.count) Teams without owners." | Log-it -Status General

#Pipe all teams w/o owners to the Get-NoOwnerTeamMembers function to find members for each team
$TeamMembers = $NoOwners | Get-NoOwnerTeamMembers

#Create Report
if($TeamMembers.Results){
    $Props = @("TeamName","Email","Members","CreationDate")
    $TeamMembersReport = $TeamMembers.Results | Sort TeamName | New-Report -Topic "Teams Without Owners" -TableName "Members of Teams Without Owners" -Properties $Props -Count $TeamMembers.Results.Count
    $TeamMembersReport.Table | Sort TeamName | Export-CSV -NoTypeInformation -Path $TeamsNoOwnersCSV -Append
}
if($TeamMembersReport.HTML){
    $TeamMembersReportHTML = $TeamMembersReport.HTML
}

#Email Report
if($TeamMembersReportHTML){$TeamMembersReportHTML | Email-Report -MailTo $To -SMTP $SMTP -Attachments $TeamsNoOwnersCSV -Title "Teams Without Owners" -Type Reporting}

#$TeamMembers.Results | Sort TeamName | Export-Csv -Path $TeamsNoOwnersCSV -NoTypeInformation

"End Script" | Log-it -Status General

#endregion process