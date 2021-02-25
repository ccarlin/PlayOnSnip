#
# powershell script to remove Playon tags and/or commercials
#

# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -ExecutionPolicy Bypass -File "D:\Documents\My Projects\PlayOnSnip\PlayOnRemoveCommercials.ps1"

# amount of time at beginning and end of video to trim off video to remove Playon tags
$startSkipSeconds = 5
$endSkipSeconds = 6

#number of minutes to retry looking for files set to zero for no retrying
$retryTime = 5

# where are the videos to convert
$inputFolder = "D:\Videos\PlayOn\"

# where do you want the trimmed videos to be stored
$outputFolder = "Z:\PlayOn\"

# optional compress settings
$compress = $true

# optionally delete files when done
$deleteSource = $true

###################################################################################################################
# do not change anything below here unless you know what you are doing
###################################################################################################################
$debug = $false
$deleteTempFiles = $true

# get a list of videos files to trim
$videoFilesToProcess = Get-ChildItem "$inputFolder*.mp4" -Recurse

function IsFileLocked($filePath) 
{
    Rename-Item $filePath $filePath -ErrorVariable errs -ErrorAction SilentlyContinue
    return ($errs.Count -ne 0)
}

function ProcessVideos
{

    # for each video file in the list...
    foreach ($videoFile in $videoFilesToProcess) 
    {
        $fileLocked = IsFileLocked($videoFile)
        if ($fileLocked -eq $true)
        {
            Write-Output "File: $($videoFile) is in process skipping."
            continue
        }

        # compute the file path to store the trimmed video
        $outputVideoFilePath = $videoFile.FullName.Replace($inputFolder, $outputFolder)
        $outputVideoFolderPath = $outputVideoFilePath.Replace($videoFile.Name,"")

        # if it already exists, skip it, else continue...
        if (!(Test-Path -Path $outputVideoFilePath)) 
        {
            Write-Output "`r`n- Processing: $($videoFile.FullName)"
        
            # if the output video folder does not exist, create it.
            if (!(Test-Path -Path $outputVideoFolderPath )) { New-Item -ItemType directory -Path $outputVideoFolderPath | Out-Null }
        
            # get a table of the chapters in the video
            $chaptersInVideo = & ffprobe -loglevel panic $videoFile -show_chapters -print_format json | ConvertFrom-Json
            $tblChaptersInVideo = $chaptersInVideo.chapters
            $numChaptersInVideo = $tblChaptersInVideo.Length

            if ($debug) { Write-Output "-- Number of chapters: $($numChaptersInVideo)." }

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
                    if ($chapterTitle -ne "Advertisement" -OR $durationSeconds -gt 600) 
                    {
                        # compute the output filename for this chapter
                        $outputChapterFile = ($tmpFilePathStart + "_{0:00}" -f $chapterId + "_" + $chapterTitle + $videoFile.Extension)

                        # trim off the start and end Playon tags
                        if ($chapterCount -eq 1) 
                        {
                            $startSeconds = $startSeconds + $startSkipSeconds
                            if ($debug) { Write-Output "--- trimming $($startSkipSeconds) seconds off start for Playon tag." }
                        }
                        elseif ($chapterCount -eq $numChaptersInVideo) 
                        {
                            $endSeconds = $endSeconds - $endSkipSeconds
                            if ($debug) { Write-Output "--- trimming $($endSkipSeconds) seconds off end for Playon tag." }
                        }

                        # compute the duration of the clip
                        $durationSeconds = $endSeconds - $startSeconds

                        # copy the clipped chapter video to the output temporary video file path
                        if ($debug) { Write-Output "-- $($startSeconds)-$($endSeconds) => $($durationSeconds): creating chapter temp file: $($outputChapterFile)." }
                        ffmpeg -loglevel panic -ss $startSeconds -i $videoFile -t $durationSeconds -c copy $outputChapterFile
                    }

                    $chapterCount++
                }

                if ($debug) { Write-Output "-- building trimmed video from chapters => $($outputVideoFilePath)" }

                # load the chapter video file names to concat into the temp file
                $videosToConcat = Get-Item "$($tmpFilePathStart)_*.mp4"
                foreach ($videoToConcat in $videosToConcat) 
                {
                    "file '" + $videoToConcat.FullName + "'" | Out-File $tmpFileName -Append -encoding default
                }

                # concat all the chapter video files into the output video file
                if ($compress) { ffmpeg -loglevel panic -f concat -safe 0 -i $tmpFileName -b:v 1M -b:a 156k $outputVideoFilePath }
                else { ffmpeg -loglevel panic -f concat -safe 0 -i $tmpFileName -c copy $outputVideoFilePath }

                # cleanup temporary files
                if ($deleteTempFiles) { Remove-Item $tmpFilePathStart*.* }
            }
            else 
            { 
                # handle video files with no chapters, simply trim off the Playon tags

                # get the original end time
                $durationSecondsOrig = ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $videoFile
                if ($debug) { Write-Output "-- Original end time: $($durationSecondsOrig)." }

                # compute the new start time
                $startSeconds = $startSkipSeconds
                if ($debug) { Write-Output "--- trimming $($startSkipSeconds) seconds off start for Playon tag." }

                # compute the new duration
                $durationSeconds = $durationSecondsOrig - $startSeconds - $endSkipSeconds
                if ($debug) { Write-Output "--- trimming $($endSkipSeconds) seconds off end for Playon tag." }

                # copy the trimmed video to the output video file path and compress
                if ($debug) { Write-Output "-- $($startSeconds) => $($durationSeconds): creating trimmed video file: $($outputVideoFilePath). Compressing: $($compress). " }
                if ($compress) { ffmpeg -loglevel panic -ss $startSeconds -i $videoFile -t $durationSeconds -b:v 1M -b:a 156k $outputVideoFilePath }
                else { ffmpeg -loglevel panic -ss $startSeconds -i $videoFile -t $durationSeconds -c copy $outputVideoFilePath }
            }

            Write-Output "- Completed: $($outputVideoFilePath)"
        }
        else 
        {
            Write-Output "`r`nSKIPPING: output video already exists: $($outputVideoFilePath)"
        }

        if ($deleteSource) 
        {         
            # if the file is not locked we can delete it.
            $fileLocked = IsFileLocked($videoFile)
            if ($fileLocked -eq $false)
            {
                Write-Output "Deleting source file: $($videoFile)"
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

    if ($retryTime -eq 0) { exit }
    
    #When done sleep for a few minutes then retry process
    Write-Output "Press any key to exit, will retry in $($retryTime) minute(s), press space to retry immediately.";

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
            sleep 1
            $timerCountdown = $timerCountdown - 1
            if ($timerCountdown -eq 0) { $keyPressed = $true }
        }
    }    

} until ($continue -eq $false)


