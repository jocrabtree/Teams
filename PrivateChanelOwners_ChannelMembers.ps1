# Created By: Josh Crabtree MS SfMC
# Date: 22 May 2026

#Connect to Teams PS Module
Connect-MicrosoftTeams

#variables
$csvPath  = "<your_csv_here>"
$ExactTime = Get-Date -Format "MM-dd-yyyy_HHmm"
$LogFile = "C:\temp\TeamsPrivateChannelNoOwner-$($ExactTime).log"

# Import CSV
$TeamsAndChannelsCsv = Import-Csv -Path $csvPath

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
    if($Color -ne "Magenta"){"$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Type$Message" | Out-File $Logfile -Append}
}

"Starting Script" | Log-It -Status General

# loop through csv & get all teamIDs, team names, channelIDs and channel names from the Get-TenantPrivateChannelMigrationStatus
$TeamChannelPairs = @()

"Looping through all teamIDs & private channels" | Log-It -Status General
foreach ($row in $TeamsAndChannelsCsv) {
    if ($row.TeamID -and $row.ChannelID) {
        $team = Get-Team -GroupId $row.TeamID
        $privateChannel = Get-TeamChannel -GroupId $row.TeamID -MembershipType Private | Where-Object { $_.Id -eq $row.ChannelID }
        $TeamChannelPairs += [PSCustomObject]@{
            TeamID = $team.GroupId
            TeamName = $team.DisplayName
            Archived = $team.Archived
            ChannelId = $privateChannel.Id
            ChannelName = $privateChannel.DisplayName
        }
    }
}

#Save all results to an array
$Results = @()

# Go through each Teams & Channel Pair, get team owners, channel display name, and private channel members
foreach ($pair in $TeamChannelPairs) {
    try {
        #Get Team Owners
        $teamOwners = Get-TeamUser -GroupId $pair.TeamID -Role Owner
       "Processing Team: $($pair.TeamName) | Channel: $($pair.ChannelName)" | Log-It -Status General

        # Get channel users
        $channelUsers = Get-TeamChannelUser -GroupId $pair.TeamID -DisplayName $pair.ChannelName

        # Filter eligible users to promote & exclude any guests w/ #EXT# & ensure role is not 'owner'
        $membersToPromote = $channelUsers | Where-Object {$_.User -notlike "*#EXT#*" -and $_.Role -ne "Owner"} | Sort-object name
        $selectMembersToPromote = $membersToPromote | Select-Object -First 2

        
        # Handle 0 members to promote returned
        if (-not $selectMembersToPromote -or $selectMembersToPromote.count -eq 0){
            "No eligible members to promote - skipping" | Log-It -Status Failure

            $Results += [PSCustomObject]@{
                TeamName       = $pair.TeamName
                TeamID         = $pair.TeamID
                Archived       = $pair.Archived
                TeamOwners     = ($teamOwners.Name -join ', ')
                ChannelName    = $privateChannel.DisplayName
                ChannelID      = $privateChannel.Id
                Members        = ""
                NewOwners      = ""
                Result         = "Skipped: No eligible members to promote"
            }

            continue
        }

        $OwnerResult = @()

        #foreach ($member in $membersToPromote) {
        foreach($member in $selectMembersToPromote){

            # if already an owner, then skip
            if ($channelUsers | Where-Object {
                $_.User -eq $member.User -and $_.Role -eq "Owner"
            }) {
                "$($member.Name) already an owner - skipping" | Log-It -Status General
                continue
            }

            try {
                "Promoting $($member.Name) to owner" | Log-It -Status General

                Add-TeamChannelUser -GroupId $pair.TeamID -DisplayName $privateChannel.DisplayName -User $member.User -Role Owner -ErrorAction Stop
                "Success: $($member.Name) promoted to owner." | Log-It -Status Success
                $OwnerResult += "$($member.Name):Success"
            }
            catch {
                "Failed: $($member.Name) | $($_.Exception.Message)" | Log-It -Status Failure
                $OwnerResult += "$($member.Name):Failed"
            }
        }

        # Collect results
        $Results += [PSCustomObject]@{
            TeamName       = $pair.TeamName
            TeamID         = $pair.TeamID
            Archived       = $pair.Archived
            TeamOwners     = ($teamOwners.Name -join ', ')
            ChannelName    = $pair.ChannelName
            ChannelID      = $pair.ChannelId
            ChannelMembers        = ($membersToPromote.Name -join ', ')
            #NewOwners      = ($membersToPromote.Name -join ', ')
            NewChannelOwners       = ($selectMembersToPromote.Name -join ', ')
            Result         = ($OwnerResult -join ', ')
        }
    }
    catch {
        "Error processing TeamID $($pair.TeamID): $($_.Exception.Message)" | Log-It -Status Failure
    }
}

# Export results to a csv
$exportPath = "C:\temp\TeamsPrivateChannelOwnerUpdate_$ExactTime.csv"
$Results | Export-Csv -Path $exportPath -NoTypeInformation
"Export complete: $exportPath" | Log-It -Status General
"End Script" | Log-It -Status General
