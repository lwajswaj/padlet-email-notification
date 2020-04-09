param($Timer)

$From = $Env:APPSETTING_Sender
$Recipient = $Env:APPSETTING_Recipient -split ";"
$SmtpServer = $Env:APPSETTING_SMTPServer
$apiUser = $Env:APPSETTING_apiUser
$apiKey = ConvertTo-SecureString -String $Env:APPSETTING_apiKey -AsPlainText -Force
[pscredential] $apiCredential = New-Object System.Management.Automation.PSCredential ($apiUser, $apiKey)

$Sections = (Invoke-WebRequest -Uri https://padlet.com/wall_sections?wall_id=56471570).Content | ConvertFrom-Json
$Entries = (Invoke-WebRequest -Uri  https://padlet.com/wishes?wall_id=56471570).Content | ConvertFrom-Json
$LastHour = (Get-Date -Minute 0 -Second 0 -Millisecond 0).AddHours(-1)

$EmailBody = ""

$NewEntries = ForEach($Entry In ($Entries | Where-Object -Property content_updated_at -gt -Value $LastHour)) {
  $Wall = ($Sections | Where-Object -FilterScript {$_.id -eq $Entry.wall_section_id}).Title
  $Title = $Entry.headline
  $Body = $Entry.Body
  $Link = $Entry.permalink

  if($Entry.attachment) {
    if($Entry.attachment -like "*.jpg" -or $Entry.attachment -like "*.png") {
      $Body += "<img src=""{0}"" alt=""{1}"">" -f $Entry.attachment, $Entry.attachment.split("/")[-1]
    }
    else {
      $Body += "<a href=""{0}"">{0}</a>" -f $Entry.attachment
    }
  }

  [PSCustomObject]@{
    "Wall" = $Wall;
    "Title" = $Title;
    "Body" = $Body;
    "Link" = $Link
  }
}

ForEach($Groups In ($NewEntries | Group-Object -Property "Wall")) {
  $EmailBody += "<h1>{0}</h1>" -f $Groups.Name

  ForEach($Item In $Groups.Group) {
    $EmailBody += "<h2><a href=""{1}"">{0}</a></h2>" -f $Item.Title, $Item.Link
    $EmailBody += $Item.Body
  }
}

if($EmailBody) {
  $Subject = "JIC N° 9 DE 1 - Padlet Update - {0}" -f $LastHour.ToString("dd/MM HH:mm")
  Send-MailMessage -Body $EmailBody -BodyAsHtml -To $Recipient -Subject $Subject -SmtpServer $SmtpServer -Credential $apiCredential -From $From -Encoding utf8
}
else {
  "Sin Novedad"
}