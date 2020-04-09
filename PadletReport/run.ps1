param($Timer)

$From = $Env:APPSETTING_Sender
$Recipient = $Env:APPSETTING_Recipient -split ";"
$SmtpServer = $Env:APPSETTING_SMTPServer
$apiUser = $Env:APPSETTING_apiUser
$apiKey = ConvertTo-SecureString -String $Env:APPSETTING_apiKey -AsPlainText -Force
[pscredential] $apiCredential = New-Object System.Management.Automation.PSCredential ($apiUser, $apiKey)

$Sections = (Invoke-WebRequest -Uri https://padlet.com/wall_sections?wall_id=56471570).Content | ConvertFrom-Json
$Entries = (Invoke-WebRequest -Uri  https://padlet.com/wishes?wall_id=56471570).Content | ConvertFrom-Json
$LastHour = (Get-Date -Minute 0 -Second 0 -Millisecond 0).AddHours(-7)

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

$EmailBody = "<table style=""background-color: #F6F6F6"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
$EmailBody += "<tr>"
$EmailBody += "<td width=""25%"" align=""center"" valign=""top"" style=""font-family:Arial, Helvetica, sans-serif; font-size:2px; color:#ffffff;"">.</td>"
$EmailBody += "<td width=""50%"" align=""center"" valign=""top"">"

ForEach($Groups In ($NewEntries | Group-Object -Property "Wall")) {
  $EmailBody += "<h1>{0}</h1>" -f $Groups.Name

  ForEach($Item In $Groups.Group) {
    $EmailBody += "<div style=""background-color: white"">"
    $EmailBody += "<h3 style=""text-align: center"">{0}</h3>" -f $Item.Title
    $EmailBody += $Item.Body
    $EmailBody += "<br/><a href=""{0}"" style=""background-color:#72B6CB;border:1px solid #72B6CB;border-radius:3px;color:#ffffff;display:inline-block;font-family:sans-serif;font-size:16px;line-height:44px;text-align:center;text-decoration:none;width:150px;-webkit-text-size-adjust:none;mso-hide:all;"">Leer mas &rarr;</a>" -f $Item.Link
    $EmailBody += "</div><p>&nbsp;</p>"
  }
}

$EmailBody += "</td>"
$EmailBody += "<td width=""25%"" align=""center"" valign=""top"" style=""font-family:Arial, Helvetica, sans-serif; font-size:2px; color:#ffffff;"">.</td>"
$EmailBody += "</tr>"
$EmailBody += "</table>"

if($NewEntries.Length -gt 0) {
  "Se encontraron novedades, enviandolas por email"
  $Subject = "JIC N° 9 DE 1 - Padlet Update - {0}" -f $LastHour.ToString("dd/MM")
  Send-MailMessage -Body $EmailBody -BodyAsHtml -To $Recipient -Subject $Subject -SmtpServer $SmtpServer -Credential $apiCredential -From $From -Encoding utf8
}
else {
  "Sin Novedad"
}