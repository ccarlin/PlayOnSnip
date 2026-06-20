# PlayOnSnip Process Flows

This document outlines the processing workflows in PlayOnSnip based on the `compress` config flag and the presence of chapters in the source video.

---

## Architecture Overview

All paths follow a two-stage approach:
1. **Cut stage** — trim/cut using stream copy (lossless, preserves everything)
2. **Final stage** — either compress (re-encode with configured hardware encoder) or move directly to output

---

## No Chapters Path

**Trigger:** Source video has **zero chapters**.

```
[Source MP4]
     |
     v
ffprobe: read total duration
     |
     v
startSeconds = startSkipSeconds
endSeconds   = totalDuration - endSkipSeconds
     |
     v
[Guard: skip if trimmed duration < 5 seconds]
     |
     v
--- CUT STAGE ---
ffmpeg -ss <start> -to <end> -i <source>
  -c:v:0 copy -c:a copy -c:s copy
  -map 0:v:0 -map 0:a -map 0:s?
  -> temp_cut.mp4 (stream copy, all streams preserved)
     |
     v
--- FINAL STAGE ---
  if compress:
    ffmpeg [-hwaccel <type>] -i <temp_cut>
      -map 0:v:0? -map 0:a:0? -map 0:s?
      -c:v <encoder> -b:v <videoRate> [encoder flags]
      -c:a aac -b:a <audioRate>
      -c:s mov_text
      -> temp_compressed.mp4
    Move-Item -> [Output MP4]
  else:
    Move-Item -> [Output MP4]
     |
     v
Delete temp files (if deleteTempFiles = true)
Delete source (if deleteSource = true)
```

**Key behavior:** Single cut with stream copy, then optionally re-encoded.

---

## Chapters Path

**Trigger:** Source video has **one or more chapters**.

```
[Source MP4]
     |
     v
ffprobe: read chapter metadata
     |
     v
--- CUT STAGE ---
For each chapter:
  |-- Skip if title == "Advertisement" AND 1s <= duration <= 600s
  |-- First chapter:  start += startSkipSeconds
  |-- Last chapter:   end   -= endSkipSeconds
  |-- [Guard: skip if trimmed duration < 5 seconds]
  |-- ffmpeg -ss <start> -to <end> -i <source>
  |     -c:v:0 copy -c:a copy -c:s copy
  |     -map 0:v:0 -map 0:a -map 0:s?
  |     -> temp chapter clip (stream copy, all streams preserved)
     |
     v
  if only one clip survives:
    use directly as final temp
  else:
    Build concat file list
    ffmpeg -f concat -safe 0 -i <list> -c copy
      -> temp_concat.mp4 (lossless join)
     |
     v
--- FINAL STAGE ---
  if compress:
    ffmpeg [-hwaccel <type>] -i <final_temp>
      -map 0:v:0? -map 0:a:0? -map 0:s?
      -c:v <encoder> -b:v <videoRate> [encoder flags]
      -c:a aac -b:a <audioRate>
      -c:s mov_text
      -> temp_compressed.mp4
    Move-Item -> [Output MP4]
  else:
    Move-Item -> [Output MP4]
     |
     v
Delete temp files (if deleteTempFiles = true)
Delete source (if deleteSource = true)
```

**Key behavior:** Each chapter is stream-copied, Advertisement chapters removed, PlayOn tags trimmed from first/last chapters, clips concatenated (or used directly if one), then optionally re-encoded.

---

## Summary Table

| Chapters? | `compress` | Cut command | Final |
|---|---|---|---|
| No | `false` | Stream copy → temp | Move temp → output |
| No | `true` | Stream copy → temp | HW encode temp → output |
| Yes | `false` | Stream copy per chapter → concat / single | Move concat/single → output |
| Yes | `true` | Stream copy per chapter → concat / single | HW encode concat/single → output |

---

## Unified Cut Command

Both paths use the same ffmpeg command for cutting:

```
ffmpeg -ss <start> -to <end> -i <input>
  -c:v:0 copy -c:a copy -c:s copy
  -map 0:v:0 -map 0:a -map 0:s?
  <output>
```

- `-c:v:0 copy` — stream copy the first video stream (skips thumbnail tracks at higher indices)
- `-c:a copy` — stream copy all audio streams
- `-c:s copy` — stream copy all subtitle streams
- `-map 0:v:0` — only the first video stream (ignores PNG thumbnail streams)
- `-map 0:a` — all audio streams
- `-map 0:s?` — all subtitle streams (optional)
- `-to` (not `-t`) — specifies the end timestamp directly

---

## Compress Encoder Flags

| `hardwareAccelType` | `-hwaccel` | Video encoder | Extra flags |
|---|---|---|---|
| `nvidia` | `cuda` | `h264_nvenc` | `-a53cc 1` |
| `amd` | `d3d11va` | `h264_amf` | `-vf scale=1280:720 -r 30` |
| `intel` | `qsv` | `h264_qsv` | `-preset fast -a53cc 1` |
| `none` | *(none)* | `libx264` | *(none)* |

All compress paths re-encode audio to AAC (`-c:a aac -b:a <audioRate>`) and re-mux subtitles to MP4-compatible format (`-c:s mov_text`).

---

## Safety

- **Duration guard:** Any clip with less than 5 seconds after trimming is skipped entirely (prevents corruption and data loss from near-empty segments).
- **Stream mapping:** Using `-map 0:v:0` explicitly selects the first video stream, avoiding PNG thumbnail video streams that cause encoder failures.
- **Temp file cleanup:** All intermediate files share a common base name pattern `{tmpBase}_*.*` for reliable cleanup.
