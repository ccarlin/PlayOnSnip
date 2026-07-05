#
# powershell script to remove Playon tags and/or commercials
#

# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -ExecutionPolicy Bypass -File "D:\Documents\My Projects\PlayOnSnip\PlayOnRemoveCommercials.ps1"
# Optional: Pass configuration file as parameter: -ConfigFile "C:\path\to\config.json"

# Accept optional config file parameter
param(
    [string]$ConfigFile = ""
)

# Function to load configuration from JSON file
function LoadConfiguration([string]$configPath) {
    if ([string]::IsNullOrEmpty($configPath)) {
        $configPath = Join-Path -Path $PSScriptRoot -ChildPath "PlayOnConfig.json"
    }
    
    if (Test-Path $configPath) {
        try {
            $config = Get-Content $configPath -Raw | ConvertFrom-Json
            return $config
        }
        catch {
            Write-Host "Error loading configuration file: $_"
            exit 1
        }
    }
    else {
        Write-Host "Configuration file not found: $configPath"
        exit 1
    }
}

# Load configuration
$config = LoadConfiguration $ConfigFile

# Apply configuration values to script variables
$startSkipSeconds = $config.startSkipSeconds
$endSkipSeconds = $config.endSkipSeconds
$retryTime = $config.retryTime
$autoExit = $config.autoExit
$inputFolder = $config.inputFolder
$outputMovies = $config.outputMovies
$outputTelevision = $config.outputTelevision
$movieLength = $config.movieLength
$compress = $config.compress
$videoRate = $config.videoRate
$audioRate = $config.audioRate
$hardwareAccelType = $config.hardwareAccelType
$deleteSource = $config.deleteSource
$logFileLocation = $config.logFileLocation
$debug = $config.debug
$deleteTempFiles = $config.deleteTempFiles
$runLowerPriority = $config.runLowerPriority
$useHevc = $config.useHevc


###################################################################################################################
# Begin Code
###################################################################################################################

#Run at a lower priority if running compression
if ($runLowerPriority -eq $true -And $compress -eq $true) {
    $process = Get-Process -Id $pid
    $process.PriorityClass = 'BelowNormal' 
}

# Universal message hanadler
function MessageLog($message, $msgType) {
    #check if the message is only for debugging purposes
    if ($msgType -eq "debug") {
        #check if we are set to view debug messages if not exit function
        if ($debug -ne $true) {
            return
        }
    }
    
    $timeStamp = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    $logstring = "$($timeStamp) - $($message)"

    if ($logFileLocation -eq "Console" -or $msgType -eq "console") { Write-Host $logstring }
    else { Add-content $logFileLocation -value $logstring }
}

function IsFileLocked($filePath) {    
    Rename-Item $filePath $filePath -ErrorVariable errs -ErrorAction SilentlyContinue
    return ($errs.Count -ne 0)
}

#check if the video is a movie soley by total length of the video
function isMovie($filePath) {
    #If there is no movie length then everything goes to the one directory
    if ($movieLength -eq 0) { return $true }

    #Check for regex match for TV Show s\d\de\d\d
    if ($filePath -match 's\d\de\d\d') { return $false }

    #If we don't have the season episode format check by length
    $videoLength = & ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $filePath
    if ($videoLength -gt $movieLength) { return $true }
    else { return $false }
}

function ProcessVideos {
    # get a list of videos files to trim
    $videoFilesToProcess = Get-ChildItem "$inputFolder*.mp4" -Recurse

    # only log the warning if there are videos to process, otherwise it just clutters the logs when there are no videos to process
    if ($videoFilesToProcess.Length -eq 0) {
        MessageLog "Video Files to Process: $($videoFilesToProcess) from folder $($inputFolder)" "debug"
    }   

    # for each video file in the list...
    foreach ($videoFile in $videoFilesToProcess) {
        $fileLocked = IsFileLocked $videoFile
        if ($fileLocked -eq $true) {
            MessageLog "File: $($videoFile) is in process skipping."
            continue
        }

        # compute the file path to store the trimmed video
        if (isMovie $videoFile -eq $true) { 
            $outputFolder = $outputMovies 
            $outputVideoFilePath = $outputMovies + $videoFile.Name
            $outputVideoFolderPath = $outputMovies 
        }
        else { 
            $outputFolder = $outputTelevision 
            MessageLog "Television show detected. Output Folder: $($outputFolder)" "debug"
            $outputVideoFilePath = $videoFile.FullName.Replace($inputFolder, $outputFolder)
            MessageLog "Replacing input folder $($inputFolder) with $($outputFolder). Output video file path: $($outputVideoFilePath)" "debug"
            $outputVideoFolderPath = $outputVideoFilePath.Replace($videoFile.Name, "")            
        }

        MessageLog "Output Folder: $($outputFolder), Output Video File Path: $($outputVideoFilePath), Output Video Folder Path: $($outputVideoFolderPath)" "debug"

        # if output already exists, skip it
        if (Test-Path -Path $outputVideoFilePath) {
            MessageLog "SKIPPING: output video already exists: $($outputVideoFilePath)"
            if ($deleteSource) {         
                $fileLocked = IsFileLocked $videoFile
                if ($fileLocked -eq $false) {
                    MessageLog "Deleting source file: $($videoFile)"
                    Remove-Item $videoFile 
                }
            }
            continue
        }

        MessageLog "Processing: $($videoFile.FullName)"

        # ensure the output folder exists
        if (!(Test-Path -Path $outputVideoFolderPath )) { New-Item -ItemType directory -Path $outputVideoFolderPath | Out-Null }

        # create a temp base for this video's intermediate files
        $tmpBase = [System.IO.Path]::GetTempFileName()
        $tmpBasePath = $tmpBase.Substring(0, $tmpBase.LastIndexOf('.'))

        # get chapters metadata
        $chaptersInVideo = & ffprobe -loglevel panic $videoFile -show_chapters -print_format json | ConvertFrom-Json
        $tblChaptersInVideo = $chaptersInVideo.chapters
        $numChaptersInVideo = $tblChaptersInVideo.Length

        MessageLog "Number of chapters: $($numChaptersInVideo)." "debug"

        if ($numChaptersInVideo -gt 0) {
            # ============================================================
            # CHAPTERS PATH — cut each chapter, then concat
            # ============================================================
            $chapterCount = 1
            foreach ($chapter in $tblChaptersInVideo) {
                $chapterId = $chapter.id
                $chapterStart = $chapter.start / 1000.0
                $chapterEnd   = $chapter.end / 1000.0
                $chapterDuration = $chapterEnd - $chapterStart

                # skip Advertisement chapters under 10 min
                $chapterTitle = $chapter.tags.title
                if ($chapterTitle -eq "Advertisement" -AND $chapterDuration -le 600 -AND $chapterDuration -ge 1) {
                    $chapterCount++
                    continue
                }

                # trim PlayOn tags from first and last chapters
                if ($chapterCount -eq 1) {
                    $chapterStart = $chapterStart + $startSkipSeconds
                    MessageLog "Trimming $($startSkipSeconds)s off start of first chapter" "debug"
                }
                if ($chapterCount -eq $numChaptersInVideo) {
                    $chapterEnd = $chapterEnd - $endSkipSeconds
                    MessageLog "Trimming $($endSkipSeconds)s off end of last chapter" "debug"
                }
                
                # guard against chapters smaller than 5 seconds
                $chapterLength = $chapterEnd - $chapterStart
                if ($chapterLength -lt 5) {
                    MessageLog "WARNING: chapter $($chapterId) has less than 5 seconds after trim, skipping" "debug"
                    $chapterCount++
                    continue
                }

                # cut chapter using the unified clip command
                $outputChapterFile = "$($tmpBasePath)_$("{0:00}" -f $chapterId)_$($chapterTitle).mkv"
                MessageLog "Cutting chapter $($chapterId): $($chapterStart)-$($chapterEnd) => $($outputChapterFile)" "debug"
                ffmpeg -loglevel panic -ss $chapterStart -i $videoFile -t $chapterLength -c:v:0 copy -c:a copy -c:s srt -map 0:v:0 -map 0:a -map 0:s -map_metadata -1 -avoid_negative_ts make_zero $outputChapterFile

                $chapterCount++
            }

            # concat all chapter clips (skip concat if only one clip)
            $chapterClips = Get-Item "$($tmpBasePath)_*.mkv"
            if ($chapterClips.Count -eq 1) {
                $finalTemp = $chapterClips[0].FullName
                MessageLog "Single chapter clip, skipping concat: $($finalTemp)" "debug"
            }
            else {
                foreach ($clip in $chapterClips) {
                    "file '" + $clip.FullName + "'" | Out-File $tmpBase -Append -encoding default
                }

                $tmpConcatFile = "$($tmpBasePath)_concat.mp4"
                MessageLog "Concatenating $($chapterClips.Count) chapter clips => $($tmpConcatFile)" "debug"
                # ffmpeg -loglevel panic -f concat -safe 0 -i $tmpBase -c copy -map 0 $tmpConcatFile
                ffmpeg -loglevel panic -f concat -safe 0 -i $tmpBase -c:v copy -c:a copy -c:s mov_text -map 0 $tmpConcatFile

                $finalTemp = $tmpConcatFile
            }
        }
        else {
            # ============================================================
            # NO CHAPTERS PATH — single cut with start/end trim
            # ============================================================
            $durationSecondsOrig = ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $videoFile
            MessageLog "Original duration: $($durationSecondsOrig)s" "debug"

            $clipStart = $startSkipSeconds
            $clipEnd   = $durationSecondsOrig - $endSkipSeconds
            $clipDuration = $clipEnd - $clipStart

            MessageLog "Trimming $($startSkipSeconds)s off start, $($endSkipSeconds)s off end" "debug"

            # guard against clips smaller than 5 seconds
            if ($clipDuration -lt 5) {
                MessageLog "WARNING: trimmed duration ($($clipDuration)s) is less than 5 seconds, skipping $($videoFile.Name)"
                Remove-Item $tmpBase -ErrorAction SilentlyContinue
                continue
            }

            # cut using the unified clip command
            $tmpCutFile = "$($tmpBasePath)_cut.mp4"
            MessageLog "Cutting $($clipStart)-$($clipEnd) => $($tmpCutFile)" "debug"
            ffmpeg -loglevel panic -ss $clipStart -i $videoFile -t $clipEnd -c:v:0 copy -c:a copy -c:s copy -map 0:v:0 -map 0:a -map 0:s? -map_metadata -1 -avoid_negative_ts make_zero $tmpCutFile

            $finalTemp = $tmpCutFile
        }

        # ================================================================
        # FINAL STEP — compress (if enabled) or move to output
        # ================================================================
        if ($compress) {
            $tmpCompressFile = "$($tmpBasePath)_compressed.mp4"
            MessageLog "Compressing final video to $($outputVideoFilePath)" "debug"

            if ($hardwareAccelType -eq "nvidia") {
                $nvencCodec = if ($useHevc) { "hevc_nvenc" } else { "h264_nvenc" }
                ffmpeg -hwaccel cuda -loglevel panic -i $finalTemp -map 0:v:0? -map 0:a:0? -map 0:s? -c:v $nvencCodec -preset p7 -rc vbr -b:v $videoRate -c:a aac -b:a $audioRate -c:s mov_text $tmpCompressFile
            }
            elseif ($hardwareAccelType -eq "amd") {
                $amfCodec = if ($useHevc) { "hevc_amf" } else { "h264_amf" }
                ffmpeg -hwaccel d3d11va -loglevel panic -i $finalTemp -map 0:v:0? -map 0:a:0? -map 0:s? -c:v $amfCodec -quality quality -rc vbr_peak -qp_i 18 -qp_p 20 -b:v $videoRate -c:a aac -b:a $audioRate -c:s mov_text $tmpCompressFile
            }
            elseif ($hardwareAccelType -eq "intel") {
                $qsvCodec = if ($useHevc) { "hevc_qsv" } else { "h264_qsv" }
                ffmpeg -hwaccel qsv -loglevel panic -i $finalTemp -map 0:v:0? -map 0:a:0? -map 0:s? -c:v $qsvCodec -b:v $videoRate -c:a aac -b:a $audioRate -c:s mov_text $tmpCompressFile
            }
            else {
                if ($useHevc) {
                    ffmpeg -loglevel panic -i $finalTemp -map 0:v:0? -map 0:a:0? -map 0:s? -c:v libx265 -preset medium -crf 26 -c:a aac -b:a $audioRate -c:s mov_text $tmpCompressFile
                }
                else {
                    ffmpeg -loglevel panic -i $finalTemp -map 0:v:0? -map 0:a:0? -map 0:s? -c:v libx264 -preset slower -crf 20 -c:a aac -b:a $audioRate -c:s mov_text $tmpCompressFile
                }
            }

            Move-Item $tmpCompressFile $outputVideoFilePath -Force
            MessageLog "Moved compressed video to output: $($outputVideoFilePath)"
        }
        else {
            Move-Item $finalTemp $outputVideoFilePath -Force
            MessageLog "Moved final video to output: $($outputVideoFilePath)"
        }

        # cleanup temp files
        if ($deleteTempFiles) {
            Remove-Item "$($tmpBasePath)*.*" -ErrorAction SilentlyContinue
            Remove-Item $tmpBase -ErrorAction SilentlyContinue
        }

        MessageLog "Completed: $($outputVideoFilePath)"

        if ($deleteSource) {         
            $fileLocked = IsFileLocked $videoFile
            if ($fileLocked -eq $false) {
                MessageLog "Deleting source file: $($videoFile)"
                Remove-Item $videoFile 
            }
        }
    }
}

[System.Console]::Title = "Playon Video Processor"
$continue = $true
do {
    #run thru the process now..
    ProcessVideos

    if ($autoExit -eq $true) {
        $videoFilesToProcess = Get-ChildItem "$inputFolder*.mp4" -Recurse  
        if ($videoFilesToProcess.Length -eq 0) {
            MessageLog "No more videos to process and auto exit is set to true, exiting now."
            exit
        }           
    }

    if ($retryTime -eq 0) { exit }
    
    #When done sleep for a few minutes then retry process
    MessageLog "Press any key to exit, 'r' to reload config, will retry in $($retryTime) minute(s), press space to retry immediately." "console"

    $keyPressed = $false
    $timerCountdown = $retryTime * 60
    while ($keyPressed -eq $false) {       
        if ([console]::KeyAvailable) {        
            $keyPressed = $true
            $key = $Host.UI.RawUI.ReadKey()          
            if ($key.Character -eq ' ') {
                # Space pressed - continue immediately
            }
            elseif ($key.Character -eq 'r' -or $key.Character -eq 'R') {
                # R pressed - reload configuration
                MessageLog "Reloading configuration from file..." "console"
                $config = LoadConfiguration $ConfigFile
                $startSkipSeconds = $config.startSkipSeconds
                $endSkipSeconds = $config.endSkipSeconds
                $retryTime = $config.retryTime
                $autoExit = $config.autoExit
                $inputFolder = $config.inputFolder
                $outputMovies = $config.outputMovies
                $outputTelevision = $config.outputTelevision
                $movieLength = $config.movieLength
                $compress = $config.compress
                $videoRate = $config.videoRate
                $audioRate = $config.audioRate
                $hardwareAccelType = $config.hardwareAccelType
                $deleteSource = $config.deleteSource
                $logFileLocation = $config.logFileLocation
                $deleteTempFiles = $config.deleteTempFiles
                $runLowerPriority = $config.runLowerPriority
                $useHevc = $config.useHevc
                MessageLog "Configuration reloaded successfully." "console"
                # Restart the loop with updated config
                $keyPressed = $true
            }
            else {
                # Any other key - exit
                $continue = $false
            }
        } 
        else {
            Start-Sleep 1
            $timerCountdown = $timerCountdown - 1
            if ($timerCountdown -eq 0) { $keyPressed = $true }
        }
    }    

} until ($continue -eq $false)


