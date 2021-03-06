﻿Set-Location $PSScriptRoot

# Examples of some of the parameters available for Get-RandomRestaurant.ps1
$RestaurantParams = @{
    Count = 5
    SearchOrigin = 'Medborgarplatsen, Stockholm'
    APIKey = 'abCdeFghIjkLmnOpqRstUvxYzaB-cdEfgHijKlm'
}

# Store the restaurants for use in email or output to file
$Restaurants = .\Get-RandomRestaurant.ps1 @RestaurantParams

# Examples of the parameters available for Send-LunchEmail.ps1
$MailParameters = @{
    SMTPServer = 'smtp.office365.com'
    Credential = Get-Credential
    Port = 587
    UseSsl = $true
    To = 'test@example.com'
    From = 'test@example.com'
    Text = "Vote for today's lunch restaurant!"
    Subject = 'Lunch time!'
    Restaurants = $Restaurants
    PollText = "Vote here!"
    PollTitle = 'May the best restaurant win!'
    StrawPoll = $true
    # Most of the text parameters have "support" for HTML
    PostText = @"
    This mail was generated in the middle of Stockholm, Sweden.
    <br><br>
    Your address {mail} has been registered to receive these emails, if you would like to unsubscribe you need to talk to the lunch administrator.
"@
}

.\Send-LunchEmail.ps1 @MailParameters
#.\Save-LunchToFile.ps1 -Text $MailParameters.Text -PostText $MailParameters.PostText -Restaurants $Restaurants -Poll -PollText $MailParameters.PollText -PollTitle $MailParameters.PollTitle