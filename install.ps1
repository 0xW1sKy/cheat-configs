#Requires -RunAsAdministrator

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$documents = [Environment]::GetFolderPath('mydocuments') 
$cheatpath = Join-Path $documents "/cheat"
$configpath = Join-Path $cheatpath "/.config"
$sheetpath = Join-Path $cheatpath "/cheatsheets"
$communitypath = Join-Path $sheetpath "/community"
$securitypath = Join-Path $sheetpath "/security"
$personalpath = Join-Path $sheetpath "/personal"
$conffile = Join-Path $configpath "/conf.yml"

#Check for Cheat folders and create if they do not exist.
if(!(Test-Path $cheatpath)){ New-Item -ItemType Directory -path $cheatpath }
if(!(Test-Path $configpath)){ New-Item -ItemType Directory -path $configpath }
if(!(Test-Path $sheetpath)){ New-Item -ItemType Directory -path $sheetpath }
if(!(Test-Path $communitypath)){ New-Item -ItemType Directory -path $communitypath }
if(!(Test-Path $securitypath)){ New-Item -ItemType Directory -path $securitypath }
if(!(Test-Path $personalpath)){ New-Item -ItemType Directory -path $personalpath }

#Get Latest Release of Cheat and 'install'
invoke-restmethod -uri https://api.github.com/repos/cheat/cheat/releases/latest | `
  Select-Object -expandproperty assets | `
  Select-Object browser_download_url | `
  Where-Object {$_.browser_download_url -like "*.exe.zip"} | `
  Select-Object -expandproperty browser_download_url |`
  foreach-object{ 
      $global:lastpercentage = -1
      $global:are = New-Object System.Threading.AutoResetEvent $false
      $uri = $_
      $of = $_.split('/') | Select-Object -last 1
      
      # web client
      # (!) output is buffered to disk -> great speed
      $wc = New-Object System.Net.WebClient

      Register-ObjectEvent -InputObject $wc -EventName DownloadProgressChanged -Action {
          # (!) getting event args
          $percentage = $event.sourceEventArgs.ProgressPercentage
          if($global:lastpercentage -lt $percentage)
          {
              $global:lastpercentage = $percentage
              # stackoverflow.com/questions/3896258
              Write-Host -NoNewline "`r Downloading... $percentage%"
          }
      } > $null

      Register-ObjectEvent -InputObject $wc -EventName DownloadFileCompleted -Action {
          $global:are.Set()
          Write-Host
      } > $null

      $wc.DownloadFileAsync($uri, "$cheatpath\$of");
      # ps script runs probably in one thread only (event is reised in same thread - blocking problems)
      # $global:are.WaitOne() not work
      while(!$global:are.WaitOne(500)) {}
      }
Expand-Archive -path "$cheatpath\$of" -destinationpath "$cheatpath\extracted"
if(get-item "$cheatpath\cheat.exe"){
  remove-item "$cheatpath\cheat.exe"
}
Get-ChildItem "$cheatpath\extracted\" -recurse | `
  Where-Object {$_.name -eq 'cheat-windows-amd64.exe'} | `
  Select-Object -expandproperty fullname | `
  foreach-object { move-item $_ -Destination "$cheatpath\cheat.exe" }

#cleanup
Remove-Item "$cheatpath\$of"
Get-Item "$cheatpath\extracted" | Remove-Item -Force -recurse 
$communitytest = get-childitem $communitypath
$securitytest = get-childitem $securitypath


if($communitytest.count -gt 1){
    try{
        git -C $communitypath pull
    }
    catch { 
        write-host "Unable to Update existing community cheatsheets"
        try{ 
            $communitytest |ForEach-Object{ remove-item $_ -recurse -force}
        }
        catch{ 
            write-host "Unable to remove existing items. Moving them..."
            move-item $communitypath "$communitypath.old"
            git clone https://github.com/cheat/cheatsheets $communitypath
        }
    }
}else{
    git clone https://github.com/cheat/cheatsheets $communitypath
}

if($securitytest.count -gt 1){
    try{
        git -C $securitypath pull
    }
    catch { 
        write-host "Unable to Update existing security cheatsheets"
        try{ 
            $securitytest |ForEach-Object{ remove-item $_ -recurse -force}
        }
        catch{ 
            write-host "Unable to remove existing items. Moving them..."
            move-item $securitypath "$securitypath.old"
            git clone https://github.com/andrewjkerr/security-cheatsheets $securitypath
        }
    }
}else{
    git clone https://github.com/andrewjkerr/security-cheatsheets $securitypath
}

copy-item "$PSScriptRoot\cheatsheets\*" $personalpath

$cheatconfig = @"
---
# The editor to use with 'cheat -e <sheet>'. Defaults to `$EDITOR or `$VISUAL.
editor: code

# Should 'cheat' always colorize output?
colorize: true

# Which 'chroma' colorscheme should be applied to the output?
# Options are available here:
#   https://github.com/alecthomas/chroma/tree/master/styles
style: monokai

# Which 'chroma' "formatter" should be applied?        
# One of: "terminal", "terminal256", "terminal16m"     
formatter: terminal16m

# The paths at which cheatsheets are available. Tags associated with a cheatpath
# are automatically attached to all cheatsheets residing on that path.
#
# Whenever cheatsheets share the same title (like 'tar'), the most local
# cheatsheets (those which come later in this file) take precedent over the
# less local sheets. This allows you to create your own "overides" for
# "upstream" cheatsheets.
#
# But what if you want to view the "upstream" cheatsheets instead of your own?
# Cheatsheets may be filtered via 'cheat -t <tag>' in combination with other
# commands. So, if you want to view the 'tar' cheatsheet that is tagged as
# 'community' rather than your own, you can use: cheat tar -t community
cheatpaths:

  # Paths that come earlier are considered to be the most "global", and will
  # thus be overridden by more local cheatsheets. That being the case, you
  # should probably list community cheatsheets first.  
  #
  # Note that the paths and tags listed below are placeholders. You may freely
  # change them to suit your needs.
  #
  # Community cheatsheets must be installed separately, though you may have
  # downloaded them automatically when installing 'cheat'. If not, you may
  # download them here:
  #
  # https://github.com/cheat/cheatsheets
  #
  # Once downloaded, ensure that 'path' below points to the location at which
  # you downloaded the community cheatsheets.
  - name: community
    path: $communitypath
    tags: [ community ]
    readonly: true

  # Class Cheat Sheets will take precedence over community sheets but not personal.
  - name: security
    path: $securitypath
    tags: [ security ]
    readonly: true

  # If you have personalized cheatsheets, list them last. They will take
  # precedence over the more global cheatsheets.       
  - name: personal
    path: $personalpath
    tags: [ personal ]
    readonly: false

  # While it requires no configuration here, it's also worth noting that
  # 'cheat' will automatically append directories named '.cheat' within the
  # current working directory to the 'cheatpath'. This can be very useful if
  # you'd like to closely associate cheatsheets with, for example, a directory
  # containing source code.
  #
  # Such "directory-scoped" cheatsheets will be treated as the most "local"
  # cheatsheets, and will override less "local" cheatsheets. Likewise,
  # directory-scoped cheatsheets will always be editable ('readonly: false').
"@

Set-Content -Path $conffile -Value $cheatconfig

# Add Cheat to Path
$pathobject = [Environment]::GetEnvironmentVariable('Path', [System.EnvironmentVariableTarget]::Machine)
$newpath = '{0}{1}{2}' -f $pathobject,[IO.Path]::PathSeparator,$cheatpath

[Environment]::SetEnvironmentVariable('Path', $newpath, [System.EnvironmentVariableTarget]::Machine)
[Environment]::SetEnvironmentVariable('CHEAT_CONFIG_PATH', $conffile, [System.EnvironmentVariableTarget]::Machine)
[Environment]::SetEnvironmentVariable('Path', $([System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")))
[Environment]::SetEnvironmentVariable('CHEAT_CONFIG_PATH', $conffile)