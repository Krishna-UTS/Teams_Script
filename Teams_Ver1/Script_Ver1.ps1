#Requires -Modules MicrosoftTeams
#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$TeamName,
    
    [Parameter(Mandatory=$false)]
    [string]$GroupId
)

#region Functions
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "Info",
        [string]$Color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $Color
}

function Test-TeamsConnection {
    try {
        $null = Get-Team
        return $true
    }
    catch {
        Write-Log "Not connected to Microsoft Teams. Please run Connect-MicrosoftTeams first." "Error" "Red"
        return $false
    }
}

function Get-CsvFile {
    param([string]$InitialPath)
    
    if ($InitialPath -and (Test-Path $InitialPath)) {
        return $InitialPath
    }
    
    Write-Log "Please provide the CSV file path" "Info" "Yellow"
    Write-Log "Example: C:\Users\YourName\Documents\teams_data.csv" "Info" "Cyan"
    
    $selectedPath = Read-Host "Enter the full path to your CSV file"
    
    if (Test-Path $selectedPath) {
        Write-Log "CSV file found: $selectedPath" "Info" "Green"
        return $selectedPath
    }
    
    Write-Log "File not found: $selectedPath" "Error" "Red"
    Write-Log "Please check the path and try again" "Info" "Yellow"
    return $null
}

function Import-TeamsCsv {
    param([string]$Path)
    
    try {
        # Validation of File at Location
        if (-not (Test-Path $Path)) {
            Write-Log "CSV file not found: $Path" "Error" "Red"
            return $null
        }
        
        # Check if File is Empty
        $fileInfo = Get-Item $Path
        if ($fileInfo.Length -eq 0) {
            Write-Log "CSV file is empty: $Path" "Error" "Red"
            return $null
        }
        
        # CSV File Format Check
        $firstLines = Get-Content $Path -TotalCount 5
        Write-Log "CSV file preview:" "Info" "Yellow"
        $firstLines | ForEach-Object { Write-Host "  $_" }
        
        # Import CSV Structure
        $csv = Import-Csv -Path $Path -ErrorAction Stop
        
        # Validate CSV Structure
        if ($csv.Count -eq 0) {
            Write-Log "CSV file contains no data rows (only headers)" "Error" "Red"
            return $null
        }
        
        # Verify Column Names
        $firstRow = $csv[0]
        $requiredColumns = @('User', 'Role', 'Channel')
        $missingColumns = @()
        
        foreach ($column in $requiredColumns) {
            if (-not $firstRow.PSObject.Properties.Name.Contains($column)) {
                $missingColumns += $column
            }
        }
        
        if ($missingColumns.Count -gt 0) {
            Write-Log "Missing required columns: $($missingColumns -join ', ')" "Error" "Red"
            Write-Log "Available columns: $($firstRow.PSObject.Properties.Name -join ', ')" "Info" "Yellow"
            return $null
        }
        
        Write-Log "Successfully imported CSV file with $($csv.Count) rows" "Info" "Green"
        Write-Log "Columns found: $($firstRow.PSObject.Properties.Name -join ', ')" "Info" "Green"
        return $csv
    }
    catch {
        Write-Log "Failed to import CSV file: $_" "Error" "Red"
        Write-Log "Please check that the file is a valid CSV with proper formatting" "Info" "Yellow"
        return $null
    }
}

function Get-TeamInfo {
    param([string]$TeamName)
    
    try {
        $team = Get-Team -DisplayName $TeamName
        if ($team) {
            Write-Log "Found team: $($team.DisplayName)" "Info" "Green"
            return $team
        }
        Write-Log "Team not found: $TeamName" "Error" "Red"
        return $null
    }
    catch {
        Write-Log "Error getting team info: $_" "Error" "Red"
        return $null
    }
}

function New-TeamsChannels {
    param(
        [string]$GroupId,
        [string]$ChannelName,
        [int]$Count,
        [string]$MembershipType = "Standard"
    )
    
    $successCount = 0
    for ($i = 1; $i -le $Count; $i++) {
        
        if ($i -lt 10) {
            $formattedNumber = "0$i"
        } else {
            $formattedNumber = "$i"
        }
        $currentName = "$ChannelName$formattedNumber"
        try {
            $params = @{
                GroupId = $GroupId
                DisplayName = $currentName
            }
            if ($MembershipType -eq "Private") {
                $params.MembershipType = "Private"
            }
            
            Write-Log "Creating channel: $currentName" "Info" "Yellow"
            New-TeamChannel @params
            
            Write-Log "Created channel: $currentName" "Info" "Green"
            
            $successCount++
            
            # Small delay to prevent rate limiting
            Start-Sleep -Milliseconds 500
        }
        catch {
            Write-Log "Failed to create channel $currentName : $($_.Exception.Message)" "Error" "Red"
        }
    }
    Write-Log "Channel creation completed. Successfully created $successCount out of $Count channels." "Info" "Green"
    return $successCount
}

function Update-TeamMembers {
    param(
        [string]$GroupId,
        [array]$CsvData`
    )
    
    $teamMemberStartTime = Get-Date
    Write-Log "Fetching current team members..." "Info" "Yellow"
    $currentUsers = Get-TeamUser -GroupId $GroupId
    $csvUsers = $CsvData.User
    
    # Create lookup hashtables for faster comparison
    $currentUserLookup = @{}
    $csvUserLookup = @{}
    foreach ($user in $currentUsers) {
        $currentUserLookup[$user.User] = $user
    }
    foreach ($user in $csvUsers) {
        $csvUserLookup[$user] = $true
    }
    
    # Batch process removals
    $usersToRemove = $currentUsers | Where-Object { 
        $_.Role -eq "Member" -and -not $csvUserLookup.ContainsKey($_.User)
    }
    
    if ($usersToRemove.Count -gt 0) {
        Write-Log "Removing $($usersToRemove.Count) users..." "Info" "Yellow"
        $batchSize = Get-DynamicBatchSize -RecordCount $usersToRemove.Count
        Write-Log "Using batch size: $batchSize for removal operation" "Info" "Cyan"
        for ($i = 0; $i -lt $usersToRemove.Count; $i += $batchSize) {
            $batch = $usersToRemove | Select-Object -Skip $i -First $batchSize
            foreach ($user in $batch) {
                try {
                    Remove-TeamUser -GroupId $GroupId -User $user.User
                    Write-Log "Removed user: $($user.User)" "Info" "Yellow"
                }
                catch {
                    Write-Log "Failed to remove user $($user.User) : $_" "Error" "Red"
                }
                # Add small delay to prevent rate limiting
                Start-Sleep -Milliseconds 100
            }
            # Add delay between batches
            Start-Sleep -Seconds 2
        }
    }
    
    # Batch process additions
    $usersToAdd = $CsvData | Where-Object { -not $currentUserLookup.ContainsKey($_.User) }
    
    if ($usersToAdd.Count -gt 0) {
        Write-Log "Adding $($usersToAdd.Count) users..." "Info" "Yellow"
        $batchSize = Get-DynamicBatchSize -RecordCount $usersToAdd.Count
        Write-Log "Using batch size: $batchSize for addition operation" "Info" "Cyan"
        for ($i = 0; $i -lt $usersToAdd.Count; $i += $batchSize) {
            $batch = $usersToAdd | Select-Object -Skip $i -First $batchSize
            foreach ($user in $batch) {
                try {
                    # Normalize role to proper Teams format
                    $normalizedRole = $user.Role
                    if ($user.Role -eq "member" -or $user.Role -eq "Member") {
                        $normalizedRole = "Member"
                    } elseif ($user.Role -eq "owner" -or $user.Role -eq "Owner") {
                        $normalizedRole = "Owner"
                    } else {
                        Write-Log "Invalid role '$($user.Role)' for user $($user.User). Using 'Member' as default." "Warning" "Yellow"
                        $normalizedRole = "Member"
                    }
                    
                    Add-TeamUser -GroupId $GroupId -User $user.User -Role $normalizedRole
                    Write-Log "Added user: $($user.User) with role: $normalizedRole" "Info" "Green"
                }
                catch {
                    Write-Log "Failed to add user $($user.User) : $_" "Error" "Red"
                }
                # Add small delay to prevent rate limiting
                Start-Sleep -Milliseconds 100
            }
            # Add delay between batches
            Start-Sleep -Seconds 2
        }
    }
    
    $teamMemberEndTime = Get-Date
    $teamMemberDuration = $teamMemberEndTime - $teamMemberStartTime
    Write-Log "Team member operations completed in: $($teamMemberDuration.ToString('mm\:ss\.fff'))" "Info" "Green"
}

function Update-ChannelMembers {
    param(
        [string]$GroupId,
        [array]$CsvData
    )
    
    $channelMemberStartTime = Get-Date
    Write-Log "Fetching channels and members..." "Info" "Yellow"
    $channels = Get-TeamChannel -GroupId $GroupId -MembershipType Private
    
    # Create lookup for CSV data with role normalization
    $channelUserLookup = @{}
    foreach ($row in $CsvData) {
        if (-not $channelUserLookup.ContainsKey($row.Channel)) {
            $channelUserLookup[$row.Channel] = @{}
        }
        
        # Normalize role to proper Teams format
        $normalizedRole = $row.Role
        if ($row.Role -eq "member" -or $row.Role -eq "Member") {
            $normalizedRole = "Member"
        } elseif ($row.Role -eq "owner" -or $row.Role -eq "Owner") {
            $normalizedRole = "Owner"
        } else {
            Write-Log "Invalid role '$($row.Role)' for user $($row.User) in channel $($row.Channel). Using 'Member' as default." "Warning" "Yellow"
            $normalizedRole = "Member"
        }
        
        $channelUserLookup[$row.Channel][$row.User] = $normalizedRole
    }
    
    # Process each channel sequentially
    foreach ($channel in $channels) {
        $channelName = $channel.DisplayName
        Write-Log "Processing channel: $channelName" "Info" "Yellow"
        
        # Skip if this channel is not in our CSV data
        if (-not $channelUserLookup.ContainsKey($channelName)) {
            Write-Log "Channel $channelName not found in CSV data, skipping..." "Info" "Yellow"
            continue
        }
        
        try {
            # Get current channel members
            $channelUsers = Get-TeamChannelUser -GroupId $GroupId -DisplayName $channelName
            
            # Create lookup for current members
            $currentMemberLookup = @{}
            foreach ($user in $channelUsers) {
                $currentMemberLookup[$user.User] = $user
            }
            
            # Get CSV data for this channel
            $channelData = $channelUserLookup[$channelName]
            if ($null -eq $channelData) { continue }
            
            # Batch process removals
            $usersToRemove = $channelUsers | Where-Object { 
                $_.Role -eq "Member" -and -not $channelData.ContainsKey($_.User)
            }
            
            if ($usersToRemove.Count -gt 0) {
                Write-Log "Removing $($usersToRemove.Count) users from channel $channelName..." "Info" "Yellow"
                $batchSize = Get-DynamicBatchSize -RecordCount $usersToRemove.Count
                Write-Log "Using batch size: $batchSize for channel removal operation" "Info" "Cyan"
                for ($i = 0; $i -lt $usersToRemove.Count; $i += $batchSize) {
                    $batch = $usersToRemove | Select-Object -Skip $i -First $batchSize
                    foreach ($user in $batch) {
                        try {
                            Remove-TeamChannelUser -GroupId $GroupId -DisplayName $channelName -User $user.User
                            Write-Log "Removed user $($user.User) from channel $channelName" "Info" "Yellow"
                        }
                        catch {
                            Write-Log "Failed to remove user from channel: $_" "Error" "Red"
                        }
                        # Add small delay to prevent rate limiting
                        Start-Sleep -Milliseconds 100
                    }
                    # Add delay between batches
                    Start-Sleep -Seconds 2
                }
            }
            
            # Batch process additions
            $usersToAdd = $channelData.GetEnumerator() | Where-Object { -not $currentMemberLookup.ContainsKey($_.Key) }
            
            if ($usersToAdd.Count -gt 0) {
                Write-Log "Adding $($usersToAdd.Count) users to channel $channelName..." "Info" "Yellow"
                $batchSize = Get-DynamicBatchSize -RecordCount $usersToAdd.Count
                Write-Log "Using batch size: $batchSize for channel addition operation" "Info" "Cyan"
                for ($i = 0; $i -lt $usersToAdd.Count; $i += $batchSize) {
                    $batch = $usersToAdd | Select-Object -Skip $i -First $batchSize
                    foreach ($user in $batch) {
                        try {
                            # For private channels, we don't specify role - users are added as members by default
                            Add-TeamChannelUser -GroupId $GroupId -DisplayName $channelName -User $user.Key
                            Write-Log "Added user $($user.Key) to channel $channelName as $($user.Value)" "Info" "Green"
                        }
                        catch {
                            Write-Log "Failed to add user to channel: $_" "Error" "Red"
                        }
                        # Add small delay to prevent rate limiting
                        Start-Sleep -Milliseconds 100
                    }
                    # Add delay between batches
                    Start-Sleep -Seconds 2
                }
            }
        }
        catch {
            Write-Log "Error processing channel $channelName : $_" "Error" "Red"
        }
        
        # Add delay between channels
        Start-Sleep -Seconds 5
    }
    
    $channelMemberEndTime = Get-Date
    $channelMemberDuration = $channelMemberEndTime - $channelMemberStartTime
    Write-Log "Channel member operations completed in: $($channelMemberDuration.ToString('mm\:ss\.fff'))" "Info" "Green"
}

function Get-DynamicBatchSize {
    param([int]$RecordCount)
    
    # Dynamic batch sizing based on record count
    if ($RecordCount -le 50) {
        # Small files: smaller batches for better control
        return 25
    } elseif ($RecordCount -le 200) {
        # Medium files: standard batch size
        return 50
    } elseif ($RecordCount -le 500) {
        # Large files: larger batches for efficiency
        return 75
    } elseif ($RecordCount -le 1000) {
        # Very large files: even larger batches
        return 100
    } else {
        # Extremely large files: cap at reasonable size
        return 150
    }
}
#endregion

#region Main Script
try {
    Clear-Host
    Write-Log "MS Teams Channel Provisioning" "Info" "Green"
    Write-Log "Last updated on $(Get-Date -Format 'd MMMM yyyy')" "Info" "White"
    
    # Initialize timer
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Check Teams connection
    if (-not (Test-TeamsConnection)) {
        Write-Log "Connecting to Microsoft Teams..." "Info" "Yellow"
        Connect-MicrosoftTeams
    }
    
    # Get CSV file
    $csvPath = Get-CsvFile -InitialPath $CsvPath
    if (-not $csvPath) { throw "No valid CSV file provided" }
    
    $csvData = Import-TeamsCsv -Path $csvPath
    if (-not $csvData) { throw "Failed to import CSV data" }
    
    # Get Team information
    if (-not $TeamName) {
        $TeamName = Read-Host "Enter the Teams site name"
    }
    
    $team = Get-TeamInfo -TeamName $TeamName
    if (-not $team) { throw "Team not found" }
    
    $GroupId = $team.GroupId
    
    # Validate GroupId is not empty
    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        throw "GroupId is empty or null. Team lookup may have failed."
    }
    
    Write-Log "Using GroupId: $GroupId" "Info" "Green"
    Write-Host "`n"
    Write-Host "`n"
    # Main menu
    do {
        Write-Log "=== Main Menu ===" "Info" "Yellow"
        Write-Log "1. Create Standard Channels" "Info" "Cyan"
        Write-Log "2. Create Private Channels" "Info" "Cyan"
        Write-Log "3. Create Private Channels from CSV" "Info" "Cyan"
        Write-Log "4. Update Team Members" "Info" "Cyan"
        Write-Log "5. Exit" "Info" "Cyan"
        
        $choice = Read-Host "Select an option"
        
        switch ($choice) {
            "1" {
                $channelName = Read-Host "Enter channel name prefix"
                $countInput = Read-Host "Enter number of channels"
                if ([int]::TryParse($countInput, [ref]$null)) {
                    $count = [int]$countInput
                    New-TeamsChannels -GroupId $GroupId -ChannelName $channelName -Count $count
                } else {
                    Write-Log "Invalid number format. Please enter a valid number." "Error" "Red"
                }
            }
            "2" {
                $channelName = Read-Host "Enter channel name prefix"
                $countInput = Read-Host "Enter number of channels"
                if ([int]::TryParse($countInput, [ref]$null)) {
                    $count = [int]$countInput
                    New-TeamsChannels -GroupId $GroupId -ChannelName $channelName -Count $count -MembershipType "Private"
                } else {
                    Write-Log "Invalid number format. Please enter a valid number." "Error" "Red"
                }
            }
            "3" {
                Write-Log "Creating private channels from CSV..." "Info" "Yellow"
                $option3Timer = [System.Diagnostics.Stopwatch]::StartNew()
                
                $existingChannels = (Get-TeamChannel -GroupId $GroupId).DisplayName
                Write-Log "Found $($existingChannels.Count) existing channels" "Info" "Cyan"
                $createdCount = 0
                $skippedCount = 0
                
                # Deduplicate channels from CSV to prevent processing the same channel multiple times
                $uniqueChannels = $csvData.Channel | Sort-Object -Unique
                Write-Log "Found $($uniqueChannels.Count) unique channels in CSV (out of $($csvData.Count) total rows)" "Info" "Cyan"
                
                # Process unique channels instead of all CSV rows
                foreach ($channelName in $uniqueChannels) {
                    if ($channelName -notin $existingChannels) {
                        $retryCount = 0
                        $maxRetries = 3
                        $channelCreated = $false
                        
                        do {
                            try {
                                Write-Log "Creating channel: $channelName (Attempt $($retryCount + 1))" "Info" "Yellow"
                                New-TeamChannel -GroupId $GroupId -DisplayName $channelName -MembershipType Private
                                Write-Log "Created channel: $channelName" "Info" "Green"
                                Write-Host "`n"
                                $createdCount++
                                $channelCreated = $true
                                Start-Sleep -Milliseconds 500
                            }
                            catch {
                                $retryCount++
                                Write-Log "Failed to create channel $channelName (Attempt $retryCount): $($_.Exception.Message)" "Error" "Red"
                                
                                if ($_.Exception.Message -like "*Channel name already existed*") {
                                    Write-Log "Channel $channelName appears to exist. Waiting 5 seconds before retry..." "Info" "Yellow"
                                    Start-Sleep -Seconds 5
                                    
                                    # Refresh the existing channels list
                                    try {
                                        $existingChannels = (Get-TeamChannel -GroupId $GroupId).DisplayName
                                        Write-Log "Refreshed channel list. Found $($existingChannels.Count) channels" "Info" "Cyan"
                                        
                                        if ($channelName -in $existingChannels) {
                                            Write-Log "Channel $channelName now found in existing list, skipping" "Info" "Yellow"
                                            $channelCreated = $true
                                            $skippedCount++
                                            break
                                        }
                                    }
                                    catch {
                                        Write-Log "Failed to refresh channel list: $($_.Exception.Message)" "Error" "Red"
                                    }
                                } else {
                                    # For other errors, don't retry
                                    Write-Log "Non-retryable error, stopping attempts for $channelName" "Error" "Red"
                                    break
                                }
                            }
                        } while (-not $channelCreated -and $retryCount -lt $maxRetries)
                        
                        if (-not $channelCreated) {
                            Write-Log "Failed to create channel $channelName after $maxRetries attempts" "Error" "Red"
                        }
                    } else {
                        Write-Log "Channel $channelName already exists, skipping..." "Info" "Yellow"
                        $skippedCount++
                    }
                }
                
                $option3Timer.Stop()
                
                # Final verification - check what channels actually exist
                $finalChannels = (Get-TeamChannel -GroupId $GroupId).DisplayName
                Write-Log "=== Final Channel Verification ===" "Info" "Cyan"
                Write-Log "Expected unique channels: $($uniqueChannels.Count)" "Info" "White"
                Write-Log "Actually created: $createdCount" "Info" "Green"
                Write-Log "Skipped (already existed): $skippedCount" "Info" "Yellow"
                Write-Log "Total channels now in team: $($finalChannels.Count)" "Info" "White"
                
                # Check which channels are missing
                $missingChannels = $uniqueChannels | Where-Object { $_ -notin $finalChannels }
                if ($missingChannels.Count -gt 0) {
                    Write-Log "Missing channels: $($missingChannels -join ', ')" "Error" "Red"
                } else {
                    Write-Log "All expected channels are present" "Info" "Green"
                }
                
                Write-Log "=== Option 3 Performance Metrics ===" "Info" "Cyan"
                Write-Log "Total channels processed: $($uniqueChannels.Count)" "Info" "White"
                Write-Log "New channels created: $createdCount" "Info" "Green"
                Write-Log "Channels skipped (already existed): $skippedCount" "Info" "Yellow"
                Write-Log "Execution time: $($option3Timer.Elapsed)" "Info" "Cyan"
                Write-Log "Average time per channel: $([TimeSpan]::FromMilliseconds($option3Timer.ElapsedMilliseconds / $uniqueChannels.Count))" "Info" "Cyan"
                Write-Log "CSV channel creation completed. Created $createdCount new channels." "Info" "Green"
            }
            "4" {
                Write-Log "Starting team and channel member updates..." "Info" "Yellow"
                $option4Timer = [System.Diagnostics.Stopwatch]::StartNew()
                
                # First, ensure all required channels exist
                Write-Log "Checking and creating missing channels..." "Info" "Yellow"
                $existingChannels = (Get-TeamChannel -GroupId $GroupId).DisplayName
                Write-Log "Found $($existingChannels.Count) existing channels" "Info" "Cyan"
                $channelsCreated = 0
                
                # Deduplicate channels from CSV to prevent processing the same channel multiple times
                $uniqueChannels = $csvData.Channel | Sort-Object -Unique
                Write-Log "Found $($uniqueChannels.Count) unique channels in CSV (out of $($csvData.Count) total rows)" "Info" "Cyan"
                
                # Process unique channels instead of all CSV rows
                foreach ($channelName in $uniqueChannels) {
                    if ($channelName -notin $existingChannels) {
                        $retryCount = 0
                        $maxRetries = 3
                        $channelCreated = $false
                        
                        do {
                            try {
                                Write-Log "Creating missing channel: $channelName (Attempt $($retryCount + 1))" "Info" "Yellow"
                                New-TeamChannel -GroupId $GroupId -DisplayName $channelName -MembershipType Private
                                Write-Log "Created channel: $channelName" "Info" "Green"
                                $channelsCreated++
                                $channelCreated = $true
                                Start-Sleep -Milliseconds 500
                            }
                            catch {
                                $retryCount++
                                Write-Log "Failed to create channel $channelName (Attempt $retryCount): $($_.Exception.Message)" "Error" "Red"
                                
                                if ($_.Exception.Message -like "*Channel name already existed*") {
                                    Write-Log "Channel $channelName appears to exist. Waiting 5 seconds before retry..." "Info" "Yellow"
                                    Start-Sleep -Seconds 5
                                    
                                    # Refresh the existing channels list
                                    try {
                                        $existingChannels = (Get-TeamChannel -GroupId $GroupId).DisplayName
                                        Write-Log "Refreshed channel list. Found $($existingChannels.Count) channels" "Info" "Cyan"
                                        
                                        if ($channelName -in $existingChannels) {
                                            Write-Log "Channel $channelName now found in existing list, skipping" "Info" "Yellow"
                                            $channelCreated = $true
                                            break
                                        }
                                    }
                                    catch {
                                        Write-Log "Failed to refresh channel list: $($_.Exception.Message)" "Error" "Red"
                                    }
                                } else {
                                    # For other errors, don't retry
                                    Write-Log "Non-retryable error, stopping attempts for $channelName" "Error" "Red"
                                    break
                                }
                            }
                        } while (-not $channelCreated -and $retryCount -lt $maxRetries)
                        
                        if (-not $channelCreated) {
                            Write-Log "Failed to create channel $channelName after $maxRetries attempts" "Error" "Red"
                        }
                    } else {
                        Write-Log "Channel $channelName already exists, skipping creation" "Info" "Yellow"
                    }
                }
                
                if ($channelsCreated -gt 0) {
                    Write-Log "Created $channelsCreated missing channels. Waiting 10 seconds for channels to fully initialize..." "Info" "Yellow"
                    Start-Sleep -Seconds 10
                }
                
                # Track team member operations
                $teamMemberTimer = [System.Diagnostics.Stopwatch]::StartNew()
                Update-TeamMembers -GroupId $GroupId -CsvData $csvData
                $teamMemberTimer.Stop()
                
                # Track channel member operations
                $channelMemberTimer = [System.Diagnostics.Stopwatch]::StartNew()
                Update-ChannelMembers -GroupId $GroupId -CsvData $csvData
                $channelMemberTimer.Stop()
                
                $option4Timer.Stop()
                
                Write-Log "=== Option 4 Performance Metrics ===" "Info" "Cyan"
                Write-Log "Total execution time: $($option4Timer.Elapsed)" "Info" "Cyan"
                Write-Log "Channels created: $channelsCreated" "Info" "White"
                Write-Log "Team member updates: $($teamMemberTimer.Elapsed)" "Info" "White"
                Write-Log "Channel member updates: $($channelMemberTimer.Elapsed)" "Info" "White"
                Write-Log "Average time per operation: $([TimeSpan]::FromMilliseconds($option4Timer.ElapsedMilliseconds / 3))" "Info" "Cyan"
            }
            "5" { break }
            default { Write-Log "Invalid option" "Error" "Red" }
        }
        
        if ($choice -ne "5") {
            Read-Host "Press Enter to continue"
        }
    } while ($choice -ne "5")
    
    # Cleanup
    if ((Read-Host "Delete CSV file? (Y/N)") -eq "Y") {
        Remove-Item -Path $csvPath -Force
        Write-Log "CSV file deleted" "Info" "Green"
    }
    
    Write-Log "Script completed successfully" "Info" "Green"
}
catch {
    Write-Log "Script failed: $_" "Error" "Red"
}
finally {
    if ($timer) {
        $timer.Stop()
        Write-Log "Execution time: $($timer.Elapsed)" "Info" "Cyan"
    }
}
#endregion