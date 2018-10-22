$RestaurantParams = @{
    Count = 3
    SearchOrigin = 'Medborgarplatsen, Stockholm'
    BlackListPath = "$PSScriptRoot\IO\Blacklist.txt"
    APIKey = 'AIzaSyCwCYzCnbLHDhcyvavMV7M-QeyVUKzTUsU'
}

Set-Location $PSScriptRoot
$Restaurants = .\Get-RandomLunchRestaurant.ps1 @RestaurantParams

$MailParameters = @{
    SMTPServer = 'smtp.test.com'
    To = 'test@example.com'
    From = 'example@test.com'
    Text = "Vote for today's lunch restaurant!"
    Subject = 'Lunch time!'
    Restaurants = $Restaurants
    PollText = "Vote here!"
    PollTitle = 'May the best restaurant win!'
    Poll = $true
    PostText = @"
    This mail was generated in the middle of Stockholm, Sweden.
    <br><br>
    Your address has been registered to receive these emails, if you would like to unsubscribe you need to talk to the lunch administrator.
"@
}

#.\Send-LunchEmail.ps1 @MailParameters

.\Save-LunchToFile.ps1 -Text $MailParameters.Text -PostText $MailParameters.PostText -Restaurants $Restaurants -Poll -PollText $MailParameters.PollText -PollTitle $MailParameters.PollTitle