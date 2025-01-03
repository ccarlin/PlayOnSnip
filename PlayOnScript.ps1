#
# powershell script to remove Playon tags and/or commercials
#

# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -ExecutionPolicy Bypass -File "D:\Documents\My Projects\PlayOnSnip\PlayOnRemoveCommercials.ps1"

# amount of time at beginning and end of video to trim off video to remove Playon tags
$startSkipSeconds = 5
$endSkipSeconds = 6

#number of minutes to retry waiting for files set to zero for no retrying
$retryTime = 5

#exit process if we are out of files to process (including in progress recordings)
$autoExit = $true

# where are the videos to convert
$inputFolder = "D:\Videos\PlayOn\"

# where do you want the trimmed videos to be stored
# if you only need one location set movieLength to 0 and put the location in outputMovies
$outputMovies = "Z:\Movies\"
$outputTelevision = "Z:\PlayOn\"
$movieLength = 4000

# optional compress settings
$compress = $true
$videoRate = "1M"
$audioRate = "156k"

# optionally delete files when done
$deleteSource = $true

# Log file location (console for screen or filepath to disk)
$logFileLocation = "d:\videos\PlayOn\LogFile.txt"
#$logFileLocation = "Console"

###################################################################################################################
# do not change anything below here unless you know what you are doing
###################################################################################################################
$debug = $false
$deleteTempFiles = $true
$runLowerPriority = $true


###################################################################################################################
# Begin Code
###################################################################################################################

#Run at a lower priority if running compression
if ($runLowerPriority -eq $true -And $compress -eq $true)
{
    $process = Get-Process -Id $pid
    $process.PriorityClass = 'BelowNormal' 
}

# Universal message hanadler
function MessageLog($message, $msgType)
{
    #check if the message is only for debugging purposes
    if ($msgType -eq "debug")
    {
        #check if we are set to view debug messages if not exit function
        if ($debug -ne $true)
        {
            return
        }
    }
    
    $timeStamp = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    $logstring = "$($timeStamp) - $($message)"

    if ($logFileLocation -eq "Console" -or $msgType -eq "console") { Write-Host $logstring }
    else { Add-content $logFileLocation -value $logstring }
}

function IsFileLocked($filePath) 
{    
    Rename-Item $filePath $filePath -ErrorVariable errs -ErrorAction SilentlyContinue
    return ($errs.Count -ne 0)
}

#check if the video is a movie soley by total length of the video
function isMovie($filePath)
{
    #If there is no movie length then everything goes to the one directory
    if ($movieLength -eq 0) { return $true }

    #Check for regex match for TV Show s\d\de\d\d
    if ($filePath -match 's\d\de\d\d') { return $false }
 
    #If we don't have the season episode format check by length
    $videoLength = & ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $filePath
    if ($videoLength -gt $movieLength) { return $true }
    else { return $false }
}

function ProcessVideos
{
    # get a list of videos files to trim
    $videoFilesToProcess = Get-ChildItem "$inputFolder*.mp4" -Recurse

    # for each video file in the list...
    foreach ($videoFile in $videoFilesToProcess) 
    {
        $fileLocked = IsFileLocked $videoFile
        if ($fileLocked -eq $true)
        {
            MessageLog "File: $($videoFile) is in process skipping."
            continue
        }

        # compute the file path to store the trimmed video
        if (isMovie $videoFile -eq $true) 
        { 
            $outputFolder = $outputMovies 
            $outputVideoFilePath = $outputMovies + $videoFile.Name
            $outputVideoFolderPath = $outputMovies 
        }
        else 
        { 
            $outputFolder = $outputTelevision 
            $outputVideoFilePath = $videoFile.FullName.Replace($inputFolder, $outputFolder)
            $outputVideoFolderPath = $outputVideoFilePath.Replace($videoFile.Name,"")
        }

        MessageLog "Output Folder: $($outputFolder), Output Video File Path: $($outputVideoFilePath), Output Video Folder Path: $($outputVideoFolderPath)" "debug"

        
        # if it already exists, skip it, else continue...
        if (!(Test-Path -Path $outputVideoFilePath)) 
        {
            MessageLog "Processing: $($videoFile.FullName)"
        
            # if the output video folder does not exist, create it.
            if (!(Test-Path -Path $outputVideoFolderPath )) { New-Item -ItemType directory -Path $outputVideoFolderPath | Out-Null }
        
            # get a table of the chapters in the video
            $chaptersInVideo = & ffprobe -loglevel panic $videoFile -show_chapters -print_format json | ConvertFrom-Json
            $tblChaptersInVideo = $chaptersInVideo.chapters
            $numChaptersInVideo = $tblChaptersInVideo.Length

            MessageLog "Number of chapters: $($numChaptersInVideo)." "debug"

            # if there are any chapters in this video, remove any "Advertisement" chapters and trim off the Playon tags
            if ($numChaptersInVideo -gt 0) 
            {
                # create a temp file to store trimmed video chapter file paths in
                $tmpFile = New-TemporaryFile
                $tmpFileName = $tmpFile.FullName
                $tmpFilePathStart = $tmpFileName.Substring(0, $tmpFileName.LastIndexOf('.'))

                $chapterCount = 1
                foreach ($chapter in $tblChaptersInVideo) 
                {
                    $chapterId = $chapter.id

                    # compute the initial duration for the chapter
                    $startSeconds = $chapter.start / 1000.0
                    $endSeconds = $chapter.end / 1000.0
                    $durationSeconds = $endSeconds - $startSeconds

                    # skip any chapters with the title of "Advertisement" unless they are longer than 5 minutes
                    #
                    # note: in some cases, chapters are mislabeled and have both video and advertisement in them
                    # so if we find any Advertisements that are longer than 5 minutes long, we include them even
                    # though the chapter has Advertisements in it so we don't miss some of the show.
                    $chapterTitle = $chapter.tags.title
                    if ($chapterTitle -ne "Advertisement" -OR $durationSeconds -gt 600 -OR $durationSeconds -lt 1) 
                    {
                        # compute the output filename for this chapter
                        $outputChapterFile = ($tmpFilePathStart + "_{0:00}" -f $chapterId + "_" + $chapterTitle + $videoFile.Extension)

                        # trim off the start and end Playon tags
                        if ($chapterCount -eq 1) 
                        {
                            $startSeconds = $startSeconds + $startSkipSeconds
                            MessageLog "Trimming $($startSkipSeconds) seconds off start for Playon tag." "debug"
                        }
                        elseif ($chapterCount -eq $numChaptersInVideo) 
                        {
                            $endSeconds = $endSeconds - $endSkipSeconds
                            MessageLog "Trimming $($endSkipSeconds) seconds off end for Playon tag." "debug"
                        }

                        # compute the duration of the clip
                        $durationSeconds = $endSeconds - $startSeconds

                        # copy the clipped chapter video to the output temporary video file path
                        MessageLog "$($startSeconds)-$($endSeconds) => $($durationSeconds): creating chapter temp file: $($outputChapterFile)." "debug"
                        # Changes from recommendation found here: https://github.com/ccarlin/PlayOnSnip/issues/1
                        # ffmpeg -loglevel panic -ss $startSeconds -i $videoFile -t $durationSeconds -c copy $outputChapterFile
                        ffmpeg -loglevel panic -ss $startSeconds -i $videoFile -t $durationSeconds -c copy -map 0:v -map 0:s -scodec mov_text -map 0:a? $outputChapterFile
                    }

                    $chapterCount++
                }

                MessageLog "Building trimmed video from chapters => $($outputVideoFilePath)" "debug"

                # load the chapter video file names to concat into the temp file
                $videosToConcat = Get-Item "$($tmpFilePathStart)_*.mp4"
                foreach ($videoToConcat in $videosToConcat) 
                {
                    "file '" + $videoToConcat.FullName + "'" | Out-File $tmpFileName -Append -encoding default
                }

                # concat all the chapter video files into the output video file
                MessageLog "Concat files $($tmpFileName) to output file $($outputVideoFilePath)"
                if ($compress) 
                { 
                    # Changes from recommendation found here: https://github.com/ccarlin/PlayOnSnip/issues/1
                    # ffmpeg -loglevel panic -f concat -safe 0 -i $tmpFileName -b:v $videoRate -b:a $audioRate -c:s mov_text $outputVideoFilePath 
                    ffmpeg -loglevel panic -f concat -safe 0 -i $tmpFileName -map 0 -b:v $videoRate -b:a $audioRate -c:s mov_text $outputVideoFilePath
                }
                else 
                { 
                    # Changes from recommendation found here: https://github.com/ccarlin/PlayOnSnip/issues/1
                    # ffmpeg -loglevel panic -f concat -safe 0 -i $tmpFileName -c copy $outputVideoFilePath 
                    # ffmpeg -loglevel panic -f concat -safe 0 -i $tmpFileName -c copy -scodec copy $outputVideoFilePath
                    ffmpeg -loglevel panic -f concat -safe 0 -i $tmpFileName -map 0 -c copy -scodec copy $outputVideoFilePath                    
                }

                # cleanup temporary files
                if ($deleteTempFiles) { Remove-Item $tmpFilePathStart*.* }
            }
            else 
            { 
                # handle video files with no chapters, simply trim off the Playon tags

                # get the original end time
                $durationSecondsOrig = ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $videoFile
                MessageLog "Original end time: $($durationSecondsOrig)." "debug"

                # compute the new start time
                $startSeconds = $startSkipSeconds
                MessageLog "Trimming $($startSkipSeconds) seconds off start for Playon tag." "debug"

                # compute the new duration
                $durationSeconds = $durationSecondsOrig - $startSeconds - $endSkipSeconds
                MessageLog "Trimming $($endSkipSeconds) seconds off end for Playon tag." "debug"

                # copy the trimmed video to the output video file path and compress
                MessageLog "$($startSeconds) => $($durationSeconds): creating trimmed video file: $($outputVideoFilePath). Compressing: $($compress). " "debug"
                if ($compress) { ffmpeg -loglevel panic -ss $startSeconds -i $videoFile -t $durationSeconds -b:v $videoRate -b:a $audioRate -c:s mov_text $outputVideoFilePath }
                else { ffmpeg -loglevel panic -ss $startSeconds -i $videoFile -t $durationSeconds -c copy -map 0:0 -map 0:1 -map 0:2 $outputVideoFilePath}
            }

           MessageLog "Completed: $($outputVideoFilePath)"
        }
        else 
        {
            MessageLog "SKIPPING: output video already exists: $($outputVideoFilePath)"
        }

        if ($deleteSource) 
        {         
            # if the file is not locked we can delete it.
            $fileLocked = IsFileLocked $videoFile
            if ($fileLocked -eq $false)
            {
                MessageLog "Deleting source file: $($videoFile)"
                Remove-Item $videoFile 
            }
        }
    }
}

$continue = $true
do
{
    #run thru the process now..
    ProcessVideos

    if ($autoExit -eq $true)
    {
        $videoFilesToProcess = Get-ChildItem "$inputFolder*.mp4" -Recurse  
        if ($videoFilesToProcess.Length -eq 0) 
        {
            MessageLog "No more videos to process and auto exit is set to true, exiting now."
            exit
        }           
    }

    if ($retryTime -eq 0) { exit }
    
    #When done sleep for a few minutes then retry process
    MessageLog "Press any key to exit, will retry in $($retryTime) minute(s), press space to retry immediately." "console"

    $keyPressed = $false
    $timerCountdown = $retryTime * 60
    while ($keyPressed -eq $false)
    {       
        if ([console]::KeyAvailable)
        {        
            $keyPressed = $true
            $key = $Host.UI.RawUI.ReadKey()          
            if ($key.Character -ne ' ') { $continue = $false }
        } 
        else
        {
            Start-Sleep 1
            $timerCountdown = $timerCountdown - 1
            if ($timerCountdown -eq 0) { $keyPressed = $true }
        }
    }    

} until ($continue -eq $false)
