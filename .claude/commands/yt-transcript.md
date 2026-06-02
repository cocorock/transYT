---
description: Download a YouTube video transcript to a formatted .docx Word file
---

You are helping the user download a YouTube transcript.

The user wants to download the transcript for this YouTube URL: $ARGUMENTS

## Steps

1. If `$ARGUMENTS` is empty, ask the user: "Please provide a YouTube URL."
   Otherwise proceed immediately.

2. Run the following command from the project root directory:
   ```
   python transyt.py "$ARGUMENTS"
   ```

3. If the command succeeds:
   - Report the saved `.docx` filename and its full path (shown in the last line of script output)
   - Mention the file is ready to open in Word or LibreOffice

4. If the command fails, report the error clearly. Common issues and responses:
   - **"No transcript available"** — this video has no captions (auto-generated or manual)
   - **"ModuleNotFoundError"** — dependencies not installed; tell the user to run `pip install -r requirements.txt`
   - **"Invalid URL"** — ask the user to confirm the URL is a valid YouTube link
   - **"Video is unavailable"** — the video may be private or deleted

Do not summarize or modify the transcript content — just run the script and report results.
