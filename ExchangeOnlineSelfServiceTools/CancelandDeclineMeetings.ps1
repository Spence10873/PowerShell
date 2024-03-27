##This script is meant to provide sample code for declining and cancelling meetings on the user's calendar using the Microsoft Graph API.
##The script uses the Invoke-RestMethod cmdlet to make HTTP requests to the Graph API and perform actions on the user's calendar.
##The script requires a valid authentication token to authenticate with the Graph API and the user's ID to identify the user's calendar.

param (
    [Parameter(Mandatory=$true)]
    [string]$userId
)

# Define the required app permissions
$requiredPermissions = "Calendars.ReadWrite", "Calendars.ReadWrite.Shared"

# Authenticate with the Graph API using the valid authentication token
# Replace <AUTH_TOKEN> with the actual authentication token
$authToken = "<AUTH_TOKEN>"
$headers = @{
    "Authorization" = "Bearer $authToken"
    "Content-Type" = "application/json"
}

# Function to handle pagination and retrieve all results
function Get-PagedResults {
    param (
        [string]$url
    )

    $results = @()

    do {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method GET
        $results += $response.value

        if ($response.'@odata.nextLink') {
            $url = $response.'@odata.nextLink'
        }
        else {
            $url = $null
        }
    } while ($url)

    return $results
}

# Get all future meetings where the user is the organizer
$organizerMeetingsUrl = "https://graph.microsoft.com/v1.0/users/$userId/events?filter=start/dateTime ge '$(Get-Date -Format s)' and isOrganizer eq true"
$organizerMeetings = Get-PagedResults -url $organizerMeetingsUrl

# Cancel all future meetings where the user is the organizer
foreach ($meeting in $organizerMeetings) {
    $cancelUrl = "https://graph.microsoft.com/v1.0/users/$userId/events/$($meeting.id)/cancel"
    Invoke-RestMethod -Uri $cancelUrl -Method POST -Headers $headers
}

# Get all future meetings where the user is an attendee
$attendeeMeetingsUrl = "https://graph.microsoft.com/v1.0/users/$userId/events?filter=start/dateTime ge '$(Get-Date -Format s)' and isOrganizer eq false and responseStatus/response eq 'accepted'"
$attendeeMeetings = Get-PagedResults -url $attendeeMeetingsUrl

# Decline all future meetings where the user is an attendee
foreach ($meeting in $attendeeMeetings) {
    $declineUrl = "https://graph.microsoft.com/v1.0/users/$userId/events/$($meeting.id)/decline"
    Invoke-RestMethod -Uri $declineUrl -Method POST -Headers $headers
}