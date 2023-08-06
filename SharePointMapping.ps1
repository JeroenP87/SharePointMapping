#Sharepointmapping
#2023.08.04 potsolutions.nl

$upn = whoami /upn
#$upn = "enter@testupn.here"

#Request odopen urls from logic app
$uri = "ADDYOURLOGICAPPURLHERE"

$postBody = @{
    upn = $upn
} | ConvertTo-Json
$response = Invoke-WebRequest -Method POST -Uri $uri -UseBasicParsing -Body $postBody -ContentType "application/json"

$sites = $response.Content.Split(',')

#wait for OneDrive to have started
Do{
    $ODStatus = Get-Process onedrive -ErrorAction SilentlyContinue
    If ($ODStatus) 
    {
        start-sleep -Seconds 10
    }
}
Until ($ODStatus)

#loop through the sites
foreach ($site in $sites) {
    if (!($site.Contains("site"))) {
        continue
    }

    #here we create a psobject and add all the queries
    $url = [uri] $site
    $ParsedQueryString = [System.Web.HttpUtility]::ParseQueryString($url.Query)

    $i = 0
    $queryParams = @()
    foreach($QueryStringObject in $ParsedQueryString){
        $queryObject = New-Object -TypeName psobject
        $queryObject | Add-Member -MemberType NoteProperty -Name Query -Value $QueryStringObject
        $queryObject | Add-Member -MemberType NoteProperty -Name Value -Value $ParsedQueryString[$i]
        $queryParams += $queryObject
        $i++
    }

    $queryParams
    $onedrivesite = ""
    $onedrivefolder = ""
    $odourl = "odopen://sync?"

    #loop through all the names and values, modify them where needed, build the odopen url tailored to the user
    foreach ($queryParam in $queryParams) {
        if ($queryParam.Query -eq "userId") { continue }
        if ($queryParam.Query -eq "isSiteAdmin") { continue }
        if ($queryParam.Query -eq "userEmail") { 
            $queryParam.Value = $upn
        }

        if ($queryParam.Query -eq "folderName") { 
            $onedrivefolder = $queryParam.Value
        }
        if ($queryParam.Query -eq "webTitle") { 
            $onedrivesite = $queryParam.Value
        }
        $odourl = $odourl + $queryParam.Query + "=" + $queryParam.Value + "&"
    }

    #check if the site has already been mapped for the current user
    $registry = Get-Item HKCU:\Software\Microsoft\OneDrive\Accounts\Business1\ScopeIdToMountPointPathCache -erroraction silentlycontinue

    foreach ($reg in $registry.Property) {
        $prop = Get-ItemProperty HKCU:\Software\Microsoft\OneDrive\Accounts\Business1\ScopeIdToMountPointPathCache -Name $reg 
            if ($prop.$($reg).EndsWith($onedrivesite + " - " + $onedrivefolder)) {
                $odourl = ""
                continue
            }
    }
    if ($odourl) {
        #execution was not skipped, new folder to map, execute the odopen url!
        Start-Process $odourl
    }
}
