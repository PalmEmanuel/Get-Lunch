﻿<#
.NOTES
    File Name : Save-LunchToFile.ps1
    Author : Emanuel Palm
    Last Edited : 2018-10-22

.Synopsis
    Saves a HTML-file in the folder where this script is run, with random restaurants and an optional poll.

.DESCRIPTION
    Saves a HTML-file in the folder where this script is run, with a list of restaurants to use as a base for what to eat for lunch.
    An optional poll can be created through API at https://www.strawpoll.me/ with the options.

    The script looks for the image .\Resources\LunchTime.png to include in the page.
    
.PARAMETER Text
    Text that will be shown as the body of the page.
    Adding HTML as part of this parameter is fine.
    
.PARAMETER PostText
    This text will show as a smaller text at the end of the page.
    Adding HTML as part of this parameter is fine.
    
.PARAMETER Restaurants
    A list of restaurants with properties Name, Rating, Website, Distance and Time where all should be strings.
    
.PARAMETER Poll
    Whether or not to create a poll to include in the page with the restaurants as options.
    
.PARAMETER PollText
    The text of the poll button.
    
.PARAMETER PollTitle
    The title of the poll.

.EXAMPLE
    .\Save-LunchToFile.ps1 -Text 'The lunch has been chosen!' -PostText 'This was generated from the middle of Stockholm!' -Restaurants $RestaurantList

    Creates a HTML-file in the folder where this script is run.
#>
[CmdletBinding(DefaultParameterSetName='Default')]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$Text,

    [Parameter()]
    [String]$PostText = '',
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [PSObject[]]$Restaurants,
    
    [Parameter(Mandatory=$true,ParameterSetName='Poll')]
    [Switch]$Poll,

    [Parameter(Mandatory=$true,ParameterSetName='Poll')]
    [ValidateNotNullOrEmpty()]
    [String]$PollText,

    [Parameter(Mandatory=$true,ParameterSetName='Poll')]
    [ValidateNotNullOrEmpty()]
    [String]$PollTitle
)

try
{
    # Dot source the Send-MailMessage script with InlineAttachments found at https://gallery.technet.microsoft.com/scriptcenter/Send-MailMessage-3a920a6d
    Test-Path "$PSScriptRoot\Dependencies\Send-MailMessage\Send-MailMessage.ps1" -ErrorAction Stop | Out-Null
    . "$PSScriptRoot\Dependencies\Send-MailMessage\Send-MailMessage.ps1"
}
catch
{
    Write-Error "Could not import dependency $PSScriptroot\Dependencies\Send-MailMessage\Send-MailMessage.ps1. Make sure the folder is created and the script is added to the folder. https://gallery.technet.microsoft.com/scriptcenter/Send-MailMessage-3a920a6d"
    exit
}

# Get random color between 0 and 255
$RGB = @()
for ($i = 0; $i -le 3-1; $i++)
{
    $RGB += Get-Random -Maximum 256
}
# Make sure it's not too bright or dark, for white and black text to work
if ((($RGB[0] -gt 200) -and ($RGB[1] -gt 200) -and ($RGB[2] -gt 200)))
{
    $RGB[(Get-Random -Maximum $RGB.Count)] -= 50
}
elseif ((($RGB[0] -lt 80) -and ($RGB[1] -lt 80) -and ($RGB[2] -lt 80)))
{
    $RGB[(Get-Random -Maximum $RGB.Count)] += 50
}

# Get Hex colors for HTML
# Loop through each RGB value, convert it to hex and then join the list and prepend a # to the result
$BackgroundColor = $(($RGB | ForEach-Object { $_.ToString('X2').ToUpper() }) -join '')
Write-Verbose "Background color: $BackgroundColor"
# Loop through each RGB value, invert it, convert it to hex and then join the list and prepend a # to the result
$InvertedColor = $(($RGB | ForEach-Object { (255 - $_) } | ForEach-Object { $_.ToString('X2').ToUpper() }) -join '')
Write-Verbose "Inverted Color: $TextColor"

# The default color for the text in the email, used for the body text for example.
$TextColor = 'FFFFFF'

if ($Poll)
{
    # Create strawpoll through API with disabled duplication checking to be able to vote several times from the same public IP
    $StrawPollURL = "https://www.strawpoll.me"
    try
    {
        $PollRequest = @{
            "title" = $($PollTitle)
            "options" = $Restaurants.Name
            "dupcheck" = 'disabled'
        }
        $PollResponse = Invoke-WebRequest -Uri "$StrawPollURL/api/v2/polls" -Method Post -Body ($PollRequest | ConvertTo-Json) -ContentType "application/json" -UseDefaultCredentials -ErrorAction Stop | ConvertFrom-Json
        
        # The response from the API call will include an ID that corresponds to our poll, we use this as link for the button.
        $PollLink = "$StrawPollURL/$($PollResponse.id)"
    }
    catch
    {
        Write-Error "Error: $($Error[0].Exception.Message)"
    }

    # The HTML for the poll "button", with text linking to the poll.
    $PollHTML = ''
    if ($PollResponse)
    {
        $PollHTML = @"
    <td align="center" valignt="top">
        <div>
        <!--[if mso]>
            <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" href="$($PollLink)" style="height:36px;v-text-anchor:middle;width:150px;" arcsize="5%" strokecolor="#$($InvertedColor)" fillcolor="#$($InvertedColor)">
            <w:anchorlock/>
            <center style="color:#$($TextColor);font-family:Helvetica, Arial,sans-serif;font-size:16px;">$($PollText)</center>
            </v:roundrect>
        <![endif]-->
        <a href="$($PollLink)" style="background-color:#$($InvertedColor);border:1px solid #$($InvertedColor);border-radius:3px;color:#$($TextColor);display:inline-block;font-family:sans-serif;font-size:16px;line-height:44px;text-align:center;text-decoration:none;width:150px;-webkit-text-size-adjust:none;mso-hide:all;">$($PollText)</a>
        </div>
    </td>
"@
    }

}

# Create HTML for the list of restaurants
# Start the HTML part by defining the amount of restaurants
$ListHTML = @"
<TH COLSPAN=$($Restaurants.Count)><hr>
"@

# Go through each restaurant and populate HTML variable
foreach ($Item in $Restaurants)
{
    # If there's a website included in the restaurant object in the list, make the text link to it
    if ($Item.Website)
    {
        $InfoHTML = @"
<td align="center" valign="middle" width="50%"><a href="$($Item.Website)"><span style="color: #$($TextColor)">$($Item.Name) &#x2605;$($Item.Rating)&#x2605;</span></a></td>
<td align="center" valign="middle" width="50%"><span style="color: #$($TextColor); text-decoration-line: underline;">$($Item.Time) ($($Item.Distance))</span></td>
"@
    }
    else
    {
        $InfoHTML = @"
<td align="center" valign="middle" width="50%"><span style="color: #$($TextColor)">$($Item.Name) &#x2605;$($Item.Rating)&#x2605;</span></td>
<td align="center" valign="middle" width="50%"><span style="color: #$($TextColor)">$($Item.Time) ($($Item.Distance))</span></td>
"@
    }

    $ListHTML += @"
<tr style="font-size: 18px; font-family: Helvetica, Arial, sans-serif; color:#$($TextColor)">
    $InfoHTML
</tr>
<TH COLSPAN=$($Restaurants.Count)><hr>
"@
}

# HTML Body of email
$HTML = @"
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <meta content="yes" name="apple-touch-fullscreen" />
  <meta name="apple-mobile-web-app-status-bar-style" content="black" />
  <meta name="format-detection" content="telephone=no" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <title>Lunch</title>
  <style type="text/css">
    a {color:#$($TextColor); padding:0; text-decoration-line: underline}
    a:link {color:#$($TextColor); text-decoration-line: underline}
    a:visited {color:#$($TextColor); text-decoration-line: underline}
    a:hover {color:#$($TextColor); text-decoration-line: underline}
    a:active {text-align: center; color:#000000; text-decoration-line: underline}
    body { width:100% !important; -webkit-text; size-adjust:100%; -ms-text-size-adjust:100%; margin:0; padding:0; background-color: #CCCCCC; }
    .ReadMsgBody { width: 100%; }
    .backgroundTable {margin:0 auto; padding:0; width:100%;!important;}
    table td {border-collapse: collapse;}
    /* Hotmail background & line height fixes */ .ExternalClass {width:100% !important;} .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height: 100%;}
    /* Image borders & formatting */ img { outline:none; -ms-interpolation-mode: bicubic;}
    a img {border:none;} /* Re-style iPhone automatic links (eg. phone numbers) */
    .applelinks a { color:#$($TextColor); }
    /* Hotmail symbol fix for mobile devices */ .ExternalClass img[class^=Emoji] { width: 10px !important; height: 10px !important; display: inline !important;}
  </style>
</head>

<body style="Margin:0;padding-top:0;padding-bottom:0;padding-right:0;padding-left:0;">
  <center class="wrapper" style="width:100%;table-layout:fixed;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;">
    <div align="center">
      <table width="100%" height="100%" align="center" cellpadding="0" border="0" cellspacing="0" style="Margin:0 auto;width:100%;">
        <tr>
          <td bgcolor="#CCCCCC" valign="top" align="center" style="background-position: top center">
            <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="40"></td>
              </tr>
            </table>
            <table width="600" border="0" align="center" bgcolor="#$($TextColor)" cellpadding="0" cellspacing="0">
              <tr>
                <td height="30" align="center" valign="top"></td>
              </tr>
              <tr>
                <td>
                  <table width="540" bgcolor="#$($BackgroundColor)" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="50" align="center" valign="top"></td>
                    </tr>
                    <tr>
                      <td height="90" align="center" valign="top">
                          <img src="$PSScriptRoot\Resources\LunchTime.png" border="0" />
                        </a>
                      </td>
                    </tr>
                    <tr>
                        <td height="100" align="center" valign="middle" style="font-size: 18px; font-family: Helvetica, Arial, sans-serif; color:#$($TextColor); padding-left: 10px; padding-right: 10px">
                        <span style="font-weight: bold; font-size: 22px color:#$($TextColor);">$($Text)</span>
                        <td height="20" align="center" valign="top"></td>
                        </td>
                    </tr>
                    </table>
                    <table width="540" bgcolor="#$($BackgroundColor)" border="0" align="center" cellpadding="0" cellspacing="0">
                        <div width="400 align="center" valign="middle">
                            $($ListHTML)
                        </div>
                    </table>
                    <table width="540" bgcolor="#$($BackgroundColor)" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                    <td height="30" align="center" valign="top"></td>
                    </tr>
                    <tr>
                        $($PollHTML)
                    </tr>
                    <td height="30" align="center" valign="top"></td>
                    </table>
                </td>
            </tr>
            <tr>
            <td height="30" align="center" valign="top"></td>
            </tr>
            </table>
            <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                <td height="27" align="center" style="font-family: Arial, Helvetica, sans-serif; color: #333333; font-weight: bold; font-size: 10px; padding: 15px">$($PostText)</td>
                </tr>
                <tr>
                <td height="60"></td>
                </tr>
            </table>
            </td>
        </tr>
        </table>
    </center>

</body>

</html>
"@

$HTML | Out-File "$PSScriptRoot\Lunch.html"