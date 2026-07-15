# Dashboard for current KPIs of access reviews in Azure Entra Identity Governance
This is a streamlit dashboard that shows the current KPIs of access reviews in Azure Entra Identity Governance. The dashboard is designed to provide a quick overview of the current state of access reviews, including the number of active and pending reviews, as well as the results of completed reviews.

## Auth
The dashboard will only be run locally and the user will manually set service principal credentials that have the correct permissions.
The dashboard then fetches a token like this:
(PowerShell Sample)
function Auth{
    param(
        [string]$clientId,
        [string]$tenantId,
        [string]$clientSecret
    )
    # Auth
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = "https://graph.microsoft.com/.default"
    }
    $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body
    $token = $tokenResponse.access_token
    return $token
}

Use the Graph REST API to get the access reviews and their results. The following endpoints will be used:
https://learn.microsoft.com/en-us/graph/api/resources/accessreviewsv2-overview?view=graph-rest-1.0

## Overview
- Show number of currently active access reviews
- Show number of pending reviews
- Group by start date, end date, and status
- show date when results will be applied
- show date when reminder emails will be sent

## Results
### Sponsor-based access reviews
Group Name = "GuestReview - $sponsorName ($month)"
- show number of not reviewed, approved, denied over all access reviews
- show access reviews with owner name that have pending reviews

### Self-Review access reviews
Group Name = "GuestReview - No Sponsor ($month)"
- show number of pending reviews, approved, denied over all access reviews