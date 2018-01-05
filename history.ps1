#Add-Type -AssemblyName System.Windows.Forms
#Add-Type -AssemblyName PresentationFramework

#hipchat API dies if you don't set TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

[string]$auth = "" #bearer token goes here
[string]$room = "2023" #2023 = CS T3
[string]$baseURI = "" #https://hipchat.domain.com/v2/room/

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.add("Authorization",$auth)

#start at current date, if running multiple queries use $csv[($csv.count - 1)].date minus one second for subsequent queries 
#note end-date parameter in $params line
$date = "2017-12-22T00:00:00" 
$params = "max-results=1000&reverse=false&start-index=0&end-date=2017-11-1T00:00:00&date=" + $date 

$uri = $baseuri + $room + "/history?" + $params
$response = Invoke-RestMethod -uri $uri -headers $headers -method get 

#wipes the CSV, disable for subsequent queries
$csv = @() 

foreach($message in $response.items){
$row = new-object system.object
    if(!$message.color){ #ignore integration messages
    $row | add-member -membertype noteproperty -name "Date" -value $message.date
    $row | add-member -membertype noteproperty -name "Name" -value $message.from.mention_name
    
        if($message.message -like '*@t3*'){
        $row | add-member -membertype noteproperty -name "T3 Ping" -value "TRUE"
        }else {$row | add-member -membertype noteproperty -name "T3 Ping" -value "FALSE"} #if/else t3
        if($message.mentions){
        $row | Add-Member -MemberType NoteProperty -Name "Mention" -value "TRUE"
        }else {$row | Add-Member -MemberType NoteProperty -Name "Mention" -Value "FALSE"} #if/else mentions
        $csv += $row
    }

    $row | add-member -membertype noteproperty -name "Message" -value $message.message
}#foreach

#export the CSV
$csv | export-csv HipChat.csv -NoTypeInformation

