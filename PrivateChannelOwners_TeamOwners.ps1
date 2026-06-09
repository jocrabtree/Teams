# Created By: Josh Crabtree MS SfMC
# Date: 22 May 2026

#Connect to Teams PS Module
Connect-MicrosoftTeams

#variables
$csvPath = "<your_csv_here>"
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

$TeamChannelPairs = $TeamChannelPairs | Sort-object TeamName

#Save all results to an array
$Results = @()

# Go through each Teams & Channel Pair, get team owners, channel display name, and private channel members
foreach ($pair in $TeamChannelPairs) {
    try {
        #Get Team Owners
        $teamOwners = Get-TeamUser -GroupId $pair.TeamID -Role Owner | Sort-object name
       "Processing Team: $($pair.TeamName), id: $($pair.TeamID) | Channel: $($pair.ChannelName), id: $($pair.ChannelID)" | Log-It -Status General

        # Get channel users
        $channelUsers = Get-TeamChannelUser -GroupId $pair.TeamID -DisplayName $pair.ChannelName

        # Get the first two team owners to promote to channel owners
        $selectOwnersToPromote = $teamOwners | Select-Object -First 2

        # Handle 0 owners to promote returned
        if (-not $selectOwnersToPromote -or $selectOwnersToPromote.count -eq 0){
            "No eligible owners to promote - skipping" | Log-It -Status Failure

            $Results += [PSCustomObject]@{
                TeamName       = $pair.TeamName
                TeamID         = $pair.TeamID
                Archived       = $pair.Archived
                TeamOwners     = ($teamOwners.Name -join ', ')
                ChannelName    = $pair.ChannelName
                ChannelID      = $pair.ChannelId
                ChannelMembers        = ""
                NewChannelOwners      = ""
                Result         = "Skipped: No eligible owners to promote"
            }

            continue
        }

        $OwnerResult = @()
        $MemberResult = @()

        foreach($owner in $selectOwnersToPromote){
            # if already an owner, then skip
            if ($channelUsers | Where-Object {
                $_.User -eq $owner.User -and $_.Role -eq "Owner"
            }) {
                "$($owner.Name) already an owner - skipping" | Log-It -Status General
                continue
            }
            try {
                "Checking if $($owner.name) is a member of channel $($pair.ChannelName)"
                 #check if team owner is a member of the private channel they are going to become an owner of
                 $existingChannelUser = $channelUsers | Where-Object { $_.User -eq $owner.User } | Select-Object -First 1
                 if($existingChannelUser){
                    "$($owner.name) is a member of channel: $($pair.ChannelName)." | Log-It -Status General
                    "Promoting $($owner.Name) to channel owner" | Log-It -Status General
                    Add-TeamChannelUser -GroupId $pair.TeamID -DisplayName $pair.ChannelName -User $owner.User -Role Owner -ErrorAction Stop
                    "Success: $($owner.Name) promoted to owner." | Log-It -Status Success
                    $OwnerResult += "$($owner.Name):Success"
                 }
                 else {
                    "$($owner.name) is not a member of channel: $($pair.ChannelName) We must add them to the channel first." | Log-It -Status General
                    Add-TeamChannelUser -GroupId $pair.TeamID -DisplayName $pair.ChannelName -User $owner.User -ErrorAction Stop
                    "Success: $($owner.Name) added to channel $($pair.ChannelName) as a member." | Log-It -Status Success
                    $MemberResult += "$($owner.name):Success"
                    
                    #grab channel members again to verify owner just added as a channel member is present
                    "Waiting five seconds before promoting $($owner.name) to an owner." | Log-It -Status General
                    Start-sleep -seconds 5
                    $channelUsers = Get-TeamChannelUser -GroupId $pair.TeamID -DisplayName $pair.ChannelName
                    $existingChannelUser = $channelUsers | Where-Object { $_.User -eq $owner.User } | Select-Object -First 1
                    if($existingChannelUser){
                        "Promoting $($owner.Name) to channel owner" | Log-It -Status General
                        Add-TeamChannelUser -GroupId $pair.TeamID -DisplayName $pair.ChannelName -User $owner.User -Role Owner -ErrorAction Stop
                        "Success: $($owner.Name) promoted to owner." | Log-It -Status Success
                        $OwnerResult += "$($owner.Name):Success"
                    }
                    else{
                        "Could not promote $($owner.Name) to owner of channel: $($pair.ChannelName). Please check this channel manually." | Log-It -Status Failure
                    }
                 }
            }
            catch {
                "Failed: $($owner.Name) | $($_.Exception.Message)" | Log-It -Status Failure
                $OwnerResult += "$($owner.Name):Failed"
            }
        }
        #Refresh channel users to get final state AFTER all changes
        $channelUsers = Get-TeamChannelUser -GroupId $pair.TeamID -DisplayName $pair.ChannelName

        # Collect results
        $Results += [PSCustomObject]@{
            TeamName       = $pair.TeamName
            TeamID         = $pair.TeamID
            Archived       = $pair.Archived
            TeamOwners     = ($teamOwners.Name -join ', ')
            ChannelName    = $pair.ChannelName
            ChannelID      = $pair.ChannelId
            ChannelMembers = ($channelUsers.Name -join ', ')
            NewChannelOwners = ($selectOwnersToPromote.Name -join ', ')
            MemberResult    = if($MemberResult){($MemberResult -join ', ')}else{"N/A"}
            OwnerResult      = ($OwnerResult -join ', ')
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
