$WorkingFolder = Split-Path -Parent $MyInvocation.MyCommand.Path

$Values = (Get-Content "$WorkingFolder\local.settings.json" | ConvertFrom-Json).Values

  $Values | Get-Member -MemberType NoteProperty `
 | Select-Object -ExpandProperty Name `
 | ForEach-Object -Process {New-Item -path Env: -Name ("APPSETTING_{0}" -f $_) -Value $Values.$_}

 if(Test-Path "$WorkingFolder\Modules") {
   Get-ChildItem -Path "$WorkingFolder\Modules" -Filter "*.psd1" -Recurse -File | ForEach-Object -Process {Import-Module -Name $_.FullName}
 }