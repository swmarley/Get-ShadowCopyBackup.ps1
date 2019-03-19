Function Get-ShadowCopyBackup{
    <#
        .SYNOPSIS
            A function used to retrieve a Shadow Copy backup of a file/folder and copy it to an alternate location.

        .DESCRIPTION
            This function retrieves a list of all available Shadow Copies on a given system and then uses parameter
            values to determine which Shadow Copy backup to use, which file/folders to extract, where to copy the
            extracted data, and if an e-mail notification should be sent. The function is able to retrieve a 
            specific Shadow Copy by formatting parameter values into a string that matches the Shadow Copy's unique 
            UNC path. This generated path is tested to make sure it is valid before the function copies its 
            contents to the desired alternate location. If an e-mail notification is desired, the specified
            recipient will be sent an e-mail containing a direct link to the alternate location the data was
            restored to.

        .PARAMETER Date
            Mandatory parameter for the desired date of the Shadow Copy backup you want to retrieve.
            Recommended format is MM/dd/yyyy

        .PARAMETER Time
            Mandatory parameter for the desired time of the Shadow Copy backup you want to retrieve. Assumes that
            you have configured Shadow Copy backups to run 3x daily [Morning = before 8am, Noon = before 1pm, 
            Evening = before 6pm].

        .PARAMETER System
            Mandatory parameter for the name of the system you want to recover data from.

        .PARAMETER Drive
            Optional parameter needed if the UNC path to the Shadow Copy backup contains a drive letter. Depending
            on your system configuration, this may or may not be necessary. It is recommended that you first check 
            the UNC path for any Shadow Copy backup on the target system before using this function.

        .PARAMETER Share
            Optional parameter needed if the UNC path to the Shadow Copy backup uses a share name, instead of a
            drive letter. Similar to the Drive parameter, it is recommended that you first check the UNC path of
            any Shadow Copy backup on the target system to determine if this is needed.

        .PARAMETER Path
            Mandatory parameter for the portion of the UNC path starting from the parent folder you wish to 
            recover from. If the full UNC path is "\\server1\share\parent\", the Path parameter should be
            "parent". If there are multiple folder levels after the share or drive, include them in the 
            Path parameter using the format "parent\subfolder".

        .PARAMETER File
            Optional parameter needed only if you want to recover an individual file. The file name should
            be formatted to include the file extension.

        .PARAMETER Destination
            Required parameter for the full UNC path to the alternate location you wish to restore data to.
            Should be written in the format "\\servername1\share\"

        .PARAMETER EmailRecipient
            Optional parameter needed if you want to send an e-mail notification to a user or group. The
            parameter value should be the full e-mail address of the desired recipient. An e-mail will
            then be sent containing a direct link to the UNC path specified in the Destination parameter.

        .EXAMPLE
            Get-ShadowCopyBackup -Date 3/7/2019 -Time Evening -System server1 -Share share1 -Path Team\User -Destination \\server2\restores -EmailRecipient user1@domain.com

        .EXAMPLE 
            Get-ShadowCopyBackup -Date 12/1/2018 -Time Noon -System server2 -Drive D -Path User\Desktop -Destination \\server3\restores

        .LINK
            GitHub Repository: https://github.com/swmarley/Get-ShadowCopyBackup.ps1
    #>
    [CmdletBinding()]

    PARAM(
        [Parameter(Mandatory=$True)]
        [string] $Date,

        [Parameter(Mandatory=$True)]
        [ValidateSet("Morning","Noon","Evening")]
        [string] $Time,

        [Parameter(Mandatory=$True)]
        [string] $System,

        [Parameter(Mandatory=$False)]
        [AllowNull()]
        [string] $Drive,

        [Parameter(Mandatory=$False)]
        [AllowNull()]
        [string] $Share,

        [Parameter(Mandatory=$True)]
        [string] $Path,

        [Parameter(Mandatory=$False)]
        [AllowNull()]
        [string] $File,

        [Parameter(Mandatory=$True)]
        [string] $Destination,
        
        [Parameter(Mandatory=$False)]
        [AllowNull()]
        [string] $EmailRecipient 
    )

    If ($Time -eq "Morning"){
        $targetTime = Get-Date $Date -Hour 8 -Minute 0 -Second 0 -Millisecond 0
    }

    Elseif($Time -eq "Noon"){
        $targetTime = Get-Date $Date -Hour 13 -Minute 0 -Second 0 -Millisecond 0
    }
    
    Elseif($Time -eq "Evening"){
        $targetTime = Get-Date $Date -Hour 18 -Minute 0 -Second 0 -Millisecond 0
    }

    $retrievedCopies = & Invoke-Command -ComputerName $System -ScriptBlock {vssadmin list shadows} | Where-Object {$_ -match "Creation"}
    $extractCopyDates = @()
    $timeSpanHash = @{}

    Foreach($copy in $retrievedCopies){
        $formattedCopy = $copy -replace ".*time:"
        $extractCopyDates += $formattedCopy
    }

    Foreach($creationDate in $extractCopyDates){
        $timeSpan = New-TimeSpan -Start $creationDate -End $targetTime | select TotalHours
        $intTimeSpan = $timeSpan.TotalHours
        $timeSpanHash.Add($creationDate, $intTimeSpan)
    }

    $closestMatch = $timeSpanHash.GetEnumerator() | Where-Object {$_.Value -ge 0} | Sort-Object {$_.Value} -Descending | select -Last 1
    
    #Adjusts shadow copy backup timestamp if it was taken prior to Daylight Savings Time (2019)
    If ((Get-Date $closestMatch.key) -lt (Get-date 3/10/2019)) {
        $dstAdjustedMatch = (Get-Date $closestMatch.key).AddHours(-1)
        $formattedMatch = (Get-Date -Date ($dstAdjustedMatch.ToUniversalTime()) -Format "yyyy.MM.dd-HH.mm.ss")
    }
    Else {
        $matchToUTC = (Get-Date -Date $closestMatch.key).ToUniversalTime()
        $formattedMatch = Get-Date -Date $matchToUTC -Format "yyyy.MM.dd-HH.mm.ss"
    }

    If($Drive){
        $folderPath = "\\$($System)\$($Drive)$\@GMT-$($formattedMatch)\$($Path)"
    }
    Elseif($Share){
        $folderPath = "\\$($System)\$($Share)\@GMT-$($formattedMatch)\$($Path)"
    }

    If(Test-Path -Path $folderPath){
        If($File){
            If(Get-ChildItem -Path $folderPath | Where-Object {$_.Name -eq "$($File)"}){
                $filePath = "$($folderPath)\$($File)"
                Copy-Item -Path $filePath -Destination "$($Destination)"
            }
        }
        Elseif($Share){
            Copy-Item -Path "$($folderPath)" -Destination "$($Destination)" -Recurse
        }
        Else{
            Write-Warning "Error: Cannot copy items from specified path."
        }
    }
    Else{
        Write-Warning "Error: The path $($folderPath) does not exist."
    }

    If ($EmailRecipient){
        $pathName = $folderPath | Out-String
        $body = "Restored items can be found at: <a href= $($Destination)>$($pathName)</a>"
        $sender = #Add sender e-mail address here
        $recipient = "$($EmailRecipient)"
        $subject = "Restored items are available from $($System) on:$($closestMatch.key)"
        $smtpServer = #Add SMTP server FQDN here

        Send-MailMessage -SmtpServer $smtpServer -to $recipient -From $sender -Body $body -BodyAsHtml -Subject $subject
    }
    Else{
        Write-Warning "Error: E-mail notification failed."
    }
}
