param($Timer)

function Get-Month{
  Param(
    [ValidateRange(1,12)]
    [int] $Month
  )

  switch ($Month) {
    1 {"Enero"}
    2 {"Febrero"}
    3 {"Marzo"}
    4 {"Abril"}
    5 {"Mayo"}
    6 {"Junio"}
    7 {"Julio"}
    8 {"Agosto"}
    9 {"Septiembre"}
    10 {"Octubre"}
    11 {"Noviembre"}
    12 {"Diciembre"}
  }
}

$TitleTemplate = @"
<tr><td bgcolor="#29D2E4">
        <div class="mktEditable" id="header">
        <table bgcolor="#13829B" border="0" cellpadding="0" cellspacing="0" width="650" class="device-width" align="center">
            <tr>
              <td style="color:#ffffff;font-weight:400;font-family:'Roboto',Arial,Sans-serif;font-size:30px;text-align:center;padding:10px 0;">
                    ##HEADER##
                </td>
            </tr>
        </table>
        </div>
</td></tr>
"@

$ArticleTemplate = @"
<tr><td>
  <div>
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="width:100% !important;text-align:center;">
        <tr>
            <td style="color:#13829B;font-weight:400;font-family:'Roboto',Arial,Sans-serif;font-size:20px;">
                ##TITLE##
            </td>
        </tr>
        <tr>
            <td height="15"></td>
        </tr>
        <tr>
            <td style="color:#848484;font-weight:400;font-family:'Roboto',Arial,Sans-serif;font-size:15px;">
                ##CONTENT##<br>
                <br>
                <table cellpadding="0" cellspacing="0" border="0" width="180" align="center" bgcolor="#13829B" style="border-right:solid 1px #0c6b80;border-left:solid 1px #0c6b80;border-top:solid 1px #0c6b80;border-bottom:solid 5px #0c6b80;border-radius:3px;">
                <tr>
                    <td style="font-size:16px;font-family:'Open Sans',Arial,Sans-serif;color:#FFFFFF;line-height:24px;text-align:center;padding:10px 20px;"><a href="##URL##" style="color:#FFFFFF;text-decoration:none;"><strong>Leer mas &rarr;</strong></a></td>
                </tr>
            </table>
            </td>
        </tr>
    </table>
    </div>
</td></tr>
<tr><td height="25"></td></tr>
"@

$SeparatorTemplate = @"
<tr>
    <td>
        <table bgcolor="#F5F5F5" border="0" cellpadding="0" cellspacing="0" width="100%" style="width:100% !important;">
            <tr>
                <td height="2" style="font-size:2px;line-height:2px;">&nbsp;</td>
            </tr>
        </table>
    </td>
</tr>
<tr><td height="25"></td></tr>
"@

$From = $Env:APPSETTING_Sender
$Recipient = $Env:APPSETTING_Recipient -split ";"
$SmtpServer = $Env:APPSETTING_SMTPServer
$apiUser = $Env:APPSETTING_apiUser
$apiKey = ConvertTo-SecureString -String $Env:APPSETTING_apiKey -AsPlainText -Force
[pscredential] $apiCredential = New-Object System.Management.Automation.PSCredential ($apiUser, $apiKey)
$FilterDate = (Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0).AddDays(-1)

if(!(Test-Path -Path "$PSScriptRoot\emailTemplate.html")) {
  throw ("One of this script dependencies is missing: {0}. Please, verify and try again" -f "$PSScriptRoot\emailTemplate.html")
}

try {
  $Sections = @((Invoke-WebRequest -Uri https://padlet.com/wall_sections?wall_id=56471570).Content | ConvertFrom-Json)
}
catch {
  $Sections = @()
}

if($Sections.Count -eq 0) {
  throw "Cannot retrieved 'Sections' from padlet"
}

try {
  $Entries = @((Invoke-WebRequest -Uri  https://padlet.com/wishes?wall_id=56471570).Content | ConvertFrom-Json)
}
catch {
  $Entries = @()
}

if($Entries.Count -eq 0) {
  throw "Cannot retrieved 'Entries' from padlet"
}

$NewEntries = ForEach($Entry In ($Entries | Where-Object -Property content_updated_at -gt -Value $FilterDate | Sort-Object -Property content_updated_at)) {
  $Wall = ($Sections | Where-Object -FilterScript {$_.id -eq $Entry.wall_section_id}).Title
  $Title = $Entry.headline
  $Body = $Entry.Body
  $Link = $Entry.permalink

  if($Entry.attachment) {
    if($Entry.attachment -like "*.jpg" -or $Entry.attachment -like "*.png") {
      $Body += "<img src=""{0}"" alt=""{1}"" width=""300"">" -f $Entry.attachment, $Entry.attachment.split("/")[-1]
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

$EmailBody = ""

ForEach($Groups In ($NewEntries | Group-Object -Property "Wall")) {
  $EmailBody += $TitleTemplate.Replace("##HEADER##", $Groups.Name)
  $EmailBody += "<tr><td>"
  $EmailBody += "<table align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""90%"" style=""widht:90% !important;margin:0 auto !important;"">"
  $EmailBody += "<tr><td height=""25""></td></tr>"


  For($i=0; $i -lt $Groups.Group.Count;$i++) {
    if($i -gt 0) {
      $EmailBody += $SeparatorTemplate
    }

    $Item = $Groups.Group[$i]
    $EmailBody += $ArticleTemplate.Replace("##TITLE##",$Item.Title).Replace("##CONTENT##",$Item.Body).Replace("##URL##",$Item.Link)
  }

  $EmailBody += "</table>"
  $EmailBody += "</td></tr>"
}

$EmailBody = (Get-Content -Path "$PSScriptRoot\emailTemplate.html" -Raw).Replace("##CONTENIDO_VA_AQUI##", $EmailBody)

if($NewEntries.Length -gt 0) {
  "News were found, sending them by email"
  $Subject = "JIC N° 9 DE 1 - Padlet Update - {0} de {1}" -f $FilterDate.Day, (Get-Month -Month $FilterDate.Month)
  Send-MailMessage -Body $EmailBody -BodyAsHtml -To $Recipient -Subject $Subject -SmtpServer $SmtpServer -Credential $apiCredential -From $From -Encoding utf8 
}
else {
  "No News"
}