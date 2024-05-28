<#
This script turns off my pc between 9pm and 7am because I forget to turn it off. 
Open task scheduler -> Create basic task -> setup wizard -> trigger at startup -> start program action -> put path
#>

# Define the shutdown time (7:00 AM)
$shutdownTimeMorning = Get-Date -Hour 7 -Minute 0 -Second 0

# Define the shutdown time (9:00 PM)
$shutdownTimeEvening = Get-Date -Hour 21 -Minute 0 -Second 0

# Define the idle and active thresholds (1 hour of inactivity, additional 30 minutes if active)
$idleThreshold = 3600  # 1 hour in seconds
$activeThreshold = 1800  # 30 minutes in seconds

# Initialize a variable to track the last active time
$lastActiveTime = $null

# Infinite loop
while ($true) {
    # Get the current time
    $currentTime = Get-Date

    # Calculate the idle time (in seconds)
    $idleTime = (New-TimeSpan -Start $lastActiveTime -End $currentTime).TotalSeconds

    # Check if the current time is within the shutdown window (9:00 PM to 7:00 AM)
    if ($currentTime.Hour -ge 21 -or $currentTime.Hour -lt 7) {
        # Check if the current time is past the shutdown time (morning)
        if ($currentTime -le $shutdownTimeMorning) {
            # Check if the computer is active
            if ($idleTime -gt $idleThreshold) {
                # Shutdown the computer
                Stop-Computer -Force
                break  # Exit the loop after shutting down
            } else {
                # Wait for additional 30 minutes
                Start-Sleep -Seconds $activeThreshold
            }
        }
        # Check if the current time is past the shutdown time (evening)
        elseif ($currentTime -ge $shutdownTimeEvening) {
            # Check if the computer is active
            if ($idleTime -gt $idleThreshold) {
                # Shutdown the computer
                Stop-Computer -Force
                break  # Exit the loop after shutting down
            } else {
                # Wait for additional 30 minutes
                Start-Sleep -Seconds $activeThreshold
            }
        } else {
            #
        }
    } else {
        #
    }

    # Check if the computer is active
    $userActivity = [System.Windows.Forms.Application]::GetLastInputInfo()
    if ($userActivity -ne 0) {
        # Update the last active time
        $lastActiveTime = $currentTime
    }

    # Delay for 5 minutes before checking again
    Start-Sleep -Seconds 300
}
