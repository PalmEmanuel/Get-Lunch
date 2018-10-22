<#
.NOTES
    File Name : Get-RandomRestaurant.ps1
    Author : Emanuel Palm
    Last Edited : 2018-10-22

.Synopsis
    Gets a specified number of random open restaurants and information about them.

.DESCRIPTION
    Specify a number of random open restaurants to get among the closest 60 restaurants to a position.
    The script returns information about the restaurants through Google's API: Name, website (if there is one), rating, distance and estimated walking time.

    To use the script you need to provide a key that has three Google APIs enabled: Distance Matrix API, Geocoding API and Places API.
    
.PARAMETER Count
    The number of restaurants to get (between 1 to 30). Returns all restaurants found if there are less restaurants than the specified number.
    Default value is 1.

.PARAMETER SearchOrigin
    The address that the search of closest restaurants will be made from.
    View this position as a middle of the circle that restaurants are found within.

.PARAMETER WalkOrigin
    If a separate address is used to determine the distance and time to restaurants.
    Examples would be if most of the restaurants were in a separate area from the office where you walk from.
    Defaults to SearchOrigin.

.PARAMETER APIKey
    The key used for the Google API calls.
    The key must have Distance Matrix API, Geocoding API and Places API enabled.

.PARAMETER Blacklist
    A string array with all the restaurants that will be excluded.

.PARAMETER BlacklistPath
    A path to a text file with all the restaurants that will be excluded.

.OUTPUTS
    A list of restaurants with Name, Rating, Website, Distance and Walking Time according to Google.

.EXAMPLE
    .\Get-RandomRestaurant.ps1 -Count 3 -SearchOrigin 'Norra Stationsgatan 67, Stockholm' -APIKey 'abCdeFghIjkLmnOpqRstUvxYzaB-cdEfgHijKlm' -Blacklist "Default Burger Place","Sushi Bar 1"

    Gets 3 random restaurants that are open right now among the 60 closest to Norra Stationsgatan 67 in Stockholm, blacklisting two restaurants called "Default Burger Place" and "Sushi Bar 1".
    The API key used for google is abCdeFghIjkLmnOpqRstUvxYzaB-cdEfgHijKlm.

.EXAMPLE
    .\Get-RandomRestaurant.ps1 -Count 5 -SearchOrigin 'Vanadisplan, Stockholm' -WalkOrigin 'Hälsingegatan 47, Stockholm' -APIKey 'abCdeFghIjkLmnOpqRstUvxYzaB-cdEfgHijKlm' -BlacklistPath 'C:\Temp\ExcludedRestaurants.txt'

    Gets 5 random restaurants that are open right now among the 60 closest to Vanadisplan in Stockholm, blacklisting any restaurants in text file "C:\Temp\ExcludedRestaurants".
    Distance to restaurants is calculated from Hälsingegatan 47 in Stockholm instead of Vanadisplan.
    The API key used for google is abCdeFghIjkLmnOpqRstUvxYzaB-cdEfgHijKlm.
#>
[Cmdletbinding(DefaultParameterSetName='Default')]
param(
    [Parameter()]
    [ValidateRange(1,30)]
    [Int]$Count = 1,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$SearchOrigin,
    
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [String]$WalkOrigin = $SearchOrigin,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$APIKey,

    [Parameter(Mandatory=$true,ParameterSetName='BlackList')]
    [String[]]$Blacklist,
    
    [Parameter(Mandatory=$true,ParameterSetName='BlackListPath')]
    [ValidatePattern('\.txt$')]
    [ValidateScript({ Test-Path $_ })]
    [String]$BlacklistPath
)

# Include system web to be able to encode the addresses to url-format
Add-Type -AssemblyName System.Web

$SearchOrigin = [System.Web.HttpUtility]::UrlEncode($SearchOrigin)
$WalkOrigin = [System.Web.HttpUtility]::UrlEncode($WalkOrigin)

# Populate Blacklist variable with content from file if parameter BlacklistPath is used
if ($PSCmdlet.ParameterSetName -eq 'BlacklistPath')
{
    try
    {
        # This throws terminating error if the blacklist file is empty, only use BlacklistPath parameter if you provide a file with content
        $Blacklist = Get-Content $BlacklistPath -ErrorAction Stop | Where { $_ -ne "" }
        
        $Output = "Read content from $BlacklistPath and found $($Blacklist.Count) restaurants to exclude: $($Blacklist -join ', ')."
        #Write-EventLog -LogName Application -Source WSH -EventId 0 -EntryType Information -Message "Get-RandomRestaurant script: $Output"
        Write-Verbose $Output
    }
    catch
    {
        Write-Error "Error reading from Blacklist file path $BlacklistPath. $($Error[0].ToString())"
        #Write-EventLog -LogName Application -Source WSH -EventId 0 -EntryType Error -Message "Get-RandomRestaurant script: Failed to read from file path $BlacklistPath. Error: $($Error[0].ToString())"
        exit
    }
}

$URL = "https://maps.googleapis.com/maps/api"

# Get Latitude and Longitude and replace commas with dots, this is needed to make API calls for the nearest restaurants
$GeoResult = Invoke-WebRequest -Method Get -ContentType "application/json" -Uri "$URL/geocode/json?address=$SearchOrigin&key=$APIKey" | ConvertFrom-Json
$Latitude = $GeoResult.results.geometry.location.lat -replace ',','.'
$Longitude = $GeoResult.results.geometry.location.lng -replace ',','.'

# Get the 20 nearest restaurants - "Page 1"
$NearbyResult = Invoke-WebRequest -Method Get -ContentType "application/json" -Uri "$URL/place/nearbysearch/json?oe=utf-8&language=sv&location=$($Latitude),$($Longitude)&type=restaurant&rankby=distance&key=$APIKey" | ConvertFrom-Json

# Save token to get next pages
$Token = $NearbyResult.next_page_token

# loop through the next_page_tokens to get 40 more restaurants, you can only get 3 "pages"
$Results = $NearbyResult.results
while ($Token)
{
    # A delay is needed for the next page token to exist on the response
    Start-Sleep -Seconds 5
    $NearbyResult = Invoke-WebRequest -Method Get -ContentType "application/json" -Uri "$URL/place/nearbysearch/json?pagetoken=$Token&key=$APIKey" | ConvertFrom-Json
    $Results += $NearbyResult.results
    $Token = $NearbyResult.next_page_token
}

# filter out blacklisted and closed restaurants, sort randomly
$RestaurantNames = $Results | Where-Object { ($_.opening_hours.open_now -eq $true) -and ($Blacklist -notcontains $_.Name) } | Select-Object -ExpandProperty name -Unique | Sort-Object { Get-Random }

Write-Verbose "Found $($RestaurantNames.Count) currently open restaurants in the area, excluding blacklisted ones."

# Empty list to contain custom PSObjects with info about restaurants
$RestaurantList = @()

# make sure we don't go out of bounds after filtering restaurants
if ($Count -gt $RestaurantNames.Count-1)
{
    $Count = $RestaurantNames.Count-1
}

# get info about filtered restaurants
for ($i = 0; $i -le $Count-1; $i++)
{
    # we don't want previous iterations to affect this one, SilentlyContinue to not get an error when they haven't been created yet
    Clear-Variable Restaurant,PlaceID,RestaurantInfo,PlaceInfo,DistanceResult,RestaurantOutput -ErrorAction SilentlyContinue

    # restaurants are already randomized so we can iterate through the list normally
    $Restaurant = $RestaurantNames[$i]
    
    # select the ID of the restaurant to send in API call
    $PlaceID = $Results | Where-Object { $_.Name -eq $Restaurant } | Select-Object -ExpandProperty place_id
    
    # Sometimes the restaurants have several places, such as a separate bar, we'll just take a guess and choose the first element in the list
    # this affects stuff such as rating and website
    if (@($PlaceID).Count -gt 1)
    {
        $PlaceID = $PlaceID[0]
    }

    # set output in url as XML - it got it to work easier than JSON
    $RestaurantInfo = Invoke-WebRequest -Method Get -ContentType "application/json" -Uri "$URL/place/details/xml?language=sv&placeid=$($PlaceID)&fields=name,rating,website&key=$APIKey"

    # extract info from xml response
    $PlaceInfo = ([xml]$RestaurantInfo.Content).PlaceDetailsResponse.result
    
    # Set proper title case format of restaurant name, some google places use ALL CAPS
    $Restaurant = (Get-Culture).TextInfo.ToTitleCase($PlaceInfo.name.ToLower())

    Write-Verbose "place_id of $($Restaurant) = $PlaceID"
    
    $DistanceResult = Invoke-WebRequest -Method Get -ContentType "application/json" -Uri "$URL/distancematrix/json?language=sv&origins=$WalkOrigin&destinations=place_id:$PlaceID&mode=walking&key=$APIKey" | ConvertFrom-Json
    
    $RestaurantOutput = New-Object PSObject -Property @{
        Name = $Restaurant
        Rating = $PlaceInfo.rating
        Website = $PlaceInfo.website
        Distance = $DistanceResult.rows.elements.distance.text
        Time = $DistanceResult.rows.elements.duration.text
    }

    Write-Verbose $RestaurantOutput

    $RestaurantList += $RestaurantOutput
}

$RestaurantList = $RestaurantList | Sort Rating

return $RestaurantList