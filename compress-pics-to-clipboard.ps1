#original written by @kenemar 8/5/22

#Make a staging folder in temp folder and make sure it's empty
New-Item -Path $env:temp\compressed-pics-staging -ItemType Directory -Force | Out-Null
Remove-Item $env:temp\compressed-pics-staging\*

#get a list of all jpegs that are less that 2 minutes old from the temp folder
$tempPics = Get-childitem $env:temp | Where-Object {($_.extension -in ".jpg",".JPG","jpeg","JPEG") -and ($_.LastWriteTime -gt (Get-Date).AddMinutes(-2))}  | %{$_.FullName}

#if the list is not empty
if($tempPics -ne $null)
{
    #copy pics to staging folder
    Copy-Item $tempPics -Destination $env:temp\compressed-pics-staging

    #make sure all matching files got copied over. if not, give error message
    $tempPics = Get-childitem $env:temp | Where-Object {($_.extension -in ".jpg",".JPG","jpeg","JPEG") -and ($_.LastWriteTime -gt (Get-Date).AddMinutes(-2))}
    $stagePics = Get-ChildItem $env:temp\compressed-pics-staging
    if(compare-Object $tempPics $stagePics)
    {
        Write-Host "Error, try again"
    } else {

        #send all pictures in staging folder to clipboard
        Set-Clipboard -Path $env:temp\compressed-pics-staging\*

        #make a string with file names of compressed pictures
        $pixList = ""
        foreach($file in $stagePics){$pixList = $pixList +  "`n" + $file.Name}
        Write-Host ("Pictures added to clipboard:" + $pixList)

        #find the name of the email window that opened (using typical window title) and close it
        $firstFileName = $tempPics[0]
        $emailWindow = Get-Process | Where-Object {$_.MainWindowTitle -like "Emailing: $firstFile*"}
        $emailWindow.CloseMainWindow() | Out-Null
    }
    
} else { Write-Host "No matching files in temp folder." }