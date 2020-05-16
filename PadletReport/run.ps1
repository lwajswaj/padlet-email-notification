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

$Environment = $Env:APPSETTING_Environment
"Environment is now = $Environment"
$ConfigurationFileName = "appSettings.json"

if($Environment) {
  $ConfigurationFileName = "appSettings.{0}.json" -f $Environment
}

"ConfigurationFileName is now = $ConfigurationFileName"

if(!(Test-Path -Path "$PSScriptRoot\$ConfigurationFileName")) {
  throw ("One of this script dependencies is missing: {0}. Please, verify and try again" -f "$PSScriptRoot\$ConfigurationFileName")
}

$Configuration = Get-Content -Path "$PSScriptRoot\$ConfigurationFileName" | ConvertFrom-Json
$SmtpServer = $Env:APPSETTING_SMTPServer
$apiUser = $Env:APPSETTING_apiUser
$apiKey = ConvertTo-SecureString -String $Env:APPSETTING_apiKey -AsPlainText -Force
[pscredential] $apiCredential = New-Object System.Management.Automation.PSCredential ($apiUser, $apiKey)
$FilterDate = (Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0).AddDays(-1)
$DescriptionRegex = New-Object System.Text.RegularExpressions.Regex("<meta name=""twitter:description"" content=""(?<Description>.+)"">")
$IsPadletUri = New-Object System.Text.RegularExpressions.Regex("http(s)?://padlet\.com/.+",[System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

"FilterDate is now = {0}" -f $FilterDate.ToString("MM/dd/yyyy")

ForEach($Padlet In ($Configuration | Where-Object -Property Enabled -eq -Value $true)) {
  "Loading Specific Variables for '{0}'" -f $Padlet.Name

  $From = $Padlet.Configuration.From
  $Recipient = $Padlet.Configuration.Recipient
  $WallId = $Padlet.Configuration.WallId
  $SeparatorColor = $Padlet.Configuration.SeparatorColor
  $TitleBackgroundColor = $Padlet.Configuration.TitleBackgroundColor
  $ButtonColor = $Padlet.Configuration.ButtonColor
  $HeaderImage = $Padlet.Configuration.HeaderImage
  $TemplateFile = $Padlet.Configuration.TemplateFile
  $Subject = "{0} - {1} de {2}" -f $Padlet.Configuration.Subject, $FilterDate.Day, (Get-Month -Month $FilterDate.Month)
  $Description = $Padlet.Configuration.Description
  $DateFilter = $Padlet.Configuration.DateFilter

  if(!$WallId) {
    throw "Wall Id not provided, cannot continue."
  }

  if(!(Test-Path -Path "$PSScriptRoot\$TemplateFile")) {
    throw ("One of this script dependencies is missing: {0}. Please, verify and try again" -f "$PSScriptRoot\$TemplateFile")
  }

  $TitleTemplate = @"
<tr><td bgcolor="#29D2E4">
        <div class="mktEditable" id="header">
        <table bgcolor="$TitleBackgroundColor" border="0" cellpadding="0" cellspacing="0" width="650" class="device-width" align="center">
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
            <td style="color:$TitleBackgroundColor;font-weight:400;font-family:'Roboto',Arial,Sans-serif;font-size:20px;">
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
                <table cellpadding="0" cellspacing="0" border="0" width="180" align="center" bgcolor="$TitleBackgroundColor" style="border-right:solid 1px $ButtonColor;border-left:solid 1px $ButtonColor;border-top:solid 1px $ButtonColor;border-bottom:solid 5px $ButtonColor;border-radius:3px;">
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
        <table bgcolor="$SeparatorColor" border="0" cellpadding="0" cellspacing="0" width="100%" style="width:100% !important;">
            <tr>
                <td height="2" style="font-size:2px;line-height:2px;">&nbsp;</td>
            </tr>
        </table>
    </td>
</tr>
<tr><td height="25"></td></tr>
"@

  if($IsPadletUri.IsMatch($Description)) {
    try {
      $Description = $DescriptionRegex.Match((invoke-webrequest -uri $Description).RawContent).Groups["Description"].Value
    }
    catch {
      $Description = ""
    }
  }

  try {
    $Sections = (Invoke-WebRequest -Uri https://padlet.com/wall_sections?wall_id=$WallId).Content | ConvertFrom-Json
  }
  catch {
    $Sections = @()
  }

  "Sections found: {0}" -f $Sections.Count

  if($Sections.Count -eq 0) {
    throw "Cannot retrieved 'Sections' from padlet"
  }

  try {
    $Entries = (Invoke-WebRequest -Uri  https://padlet.com/wishes?wall_id=$WallId).Content | ConvertFrom-Json
  }
  catch {
    $Entries = @()
  }

  "Entries found: {0}" -f $Entries.Count

  if($Entries.Count -eq 0) {
    throw "Cannot retrieved 'Entries' from padlet"
  }

  $NewEntries = @(ForEach($Entry In ($Entries | Where-Object -Property $DateFilter -gt -Value $FilterDate | Sort-Object -Property $DateFilter)) {
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
      "Wall" = "{0}|{1}" -f $Entry.wall_section_id,$Wall;
      "Title" = $Title;
      "Body" = $Body;
      "Link" = $Link
    }
  })

  "New Entries found: {0}" -f $NewEntries.Count

  if($NewEntries.Count -gt 0) {
    $EmailBody = ""

    ForEach($Groups In ($NewEntries | Group-Object -Property "Wall" | Sort-Object -Property "Name")) {
      $EmailBody += $TitleTemplate.Replace("##HEADER##", $Groups.Name.Substring($Groups.Name.IndexOf("|") + 1))
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

    $EmailBody = (Get-Content -Path "$PSScriptRoot\$TemplateFile" -Raw).Replace("##CONTENT_GOES_HERE##", $EmailBody).Replace("##DESCRIPCION##", $Description).Replace("##HEADERIMAGE##",$HeaderImage)

    "News were found, sending them by email"
    Send-MailMessage -Body $EmailBody -BodyAsHtml -To $Recipient -Subject $Subject -SmtpServer $SmtpServer -Credential $apiCredential -From $From -Encoding utf8 
  }
  else {
    "No News"
  }
}