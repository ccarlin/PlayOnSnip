# PlayOnSnip

PowerShell script to process PlayOn recordings — trim start/end tags, remove advertisement chapters, optionally compress, and route movies vs. TV shows to separate output folders.

## Features

- Trims PlayOn intro/outro tags from recorded videos
- Removes chapters marked as "Advertisement" (skips ads shorter than 10 minutes; preserves mislabeled chapters that contain actual content)
- Automatically detects movies vs. TV shows by filename regex (`s\d\de\d\d`) or total duration
- Optional video compression using ffmpeg with hardware acceleration support (NVIDIA, AMD, Intel, or software)
- Runs continuously watching a folder, or processes existing files and exits
- Deletes source files after successful processing (optional)
- Lowers process priority during compression (optional)
- Hot-reloadable configuration via `r` key press
- Configurable via external JSON file, passed as a command-line parameter

## Requirements

- [ffmpeg](https://ffmpeg.org/) installed and in your PATH (including `ffprobe`)
- PowerShell 5.0+
- (Optional) Hardware encoder drivers for GPU-accelerated compression:
  - **NVIDIA**: NVENC (included with NVIDIA drivers)
  - **AMD**: AMF (included with AMD Adrenalin drivers)
  - **Intel**: QSV (included with Intel Graphics drivers)

## Configuration

The script loads settings from a JSON configuration file. By default it looks for `PlayOnConfig.json` in the same directory as the script. A different config can be specified via the `-ConfigFile` parameter.

### `-ConfigFile` Parameter

```
powershell -ExecutionPolicy Bypass -File "PlayOnScript.ps1" -ConfigFile "C:\path\to\PlayOnConfig.json"
```

### Configuration File Reference

| Field | Type | Description |
|---|---|---|
| `startSkipSeconds` | integer | Seconds to trim from the beginning of the video |
| `endSkipSeconds` | integer | Seconds to trim from the end of the video |
| `retryTime` | integer | Minutes to wait before checking for new files again (0 = process once and exit) |
| `autoExit` | boolean | Exit automatically when no more videos are found (only in continuous mode) |
| `inputFolder` | string | Directory to watch for PlayOn recordings |
| `outputMovies` | string | Output directory for detected movies |
| `outputTelevision` | string | Output directory for detected TV shows |
| `movieLength` | integer | Minimum duration in seconds to classify a video as a movie (0 = send everything to movies folder) |
| `compress` | boolean | Enable video compression via ffmpeg |
| `videoRate` | string | Video bitrate for compression (e.g. `"1M"`, `"2M"`) |
| `audioRate` | string | Audio bitrate for compression (e.g. `"128k"`, `"156k"`) |
| `hardwareAccelType` | string | Hardware acceleration: `"nvidia"`, `"amd"`, `"intel"`, or `"none"` |
| `deleteSource` | boolean | Delete the original source file after successful processing |
| `logFileLocation` | string | Path to log file, or `"Console"` to log to console only |
| `debug` | boolean | Enable verbose debug logging |
| `deleteTempFiles` | boolean | Clean up temporary chapter files after processing |
| `runLowerPriority` | boolean | Run the script process at BelowNormal priority during compression |

### Example `PlayOnConfig.json`

```json
{
  "startSkipSeconds": 5,
  "endSkipSeconds": 6,
  "retryTime": 5,
  "autoExit": false,
  "inputFolder": "C:\\Users\\Public\\Videos\\PlayOn\\",
  "outputMovies": "D:\\Media\\Movies\\",
  "outputTelevision": "D:\\Media\\TV Shows\\",
  "movieLength": 4000,
  "compress": true,
  "videoRate": "1M",
  "audioRate": "156k",
  "hardwareAccelType": "nvidia",
  "deleteSource": true,
  "logFileLocation": "D:\\Logs\\PlayOnSnip.txt",
  "debug": false,
  "deleteTempFiles": true,
  "runLowerPriority": true
}
```

Two sample configs are included:
- `PlayOnConfigServer.json` — example for a server/NVIDIA setup
- `PlayOnConfigLocal.json` — example for a local/AMD setup

## Usage

```powershell
# Run with default config (PlayOnConfig.json in script directory)
.\PlayOnScript.ps1

# Run with a specific config file
.\PlayOnScript.ps1 -ConfigFile "D:\Configs\PlayOnConfigServer.json"
```

### Runtime Controls

While the script is running:
- **Space** — Check for new files immediately
- **R** — Reload configuration from file (applies changes without restart)
- **Any other key** — Exit the script

## How It Works

1. Scans `inputFolder` recursively for `.mp4` files
2. Skips files currently being written (lock check)
3. Determines if the video is a movie or TV show (by filename regex `s\d\de\d\d` or duration vs. `movieLength`)
4. Reads ffprobe chapter metadata
5. Splits video into chapter clips, trimming `startSkipSeconds` from the first chapter and `endSkipSeconds` from the last chapter
6. Skips chapters titled "Advertisement" (unless they're longer than 10 minutes, indicating a mislabel)
7. Concatenates remaining chapters into a single video
8. Optionally compresses with hardware-accelerated encoding
9. Moves the final video to the appropriate output folder
10. Optionally deletes the source file

## Logging

Log output is written to the file specified in `logFileLocation`, or to the console if set to `"Console"`. Enable `debug: true` for verbose chapter-by-chapter logging.
