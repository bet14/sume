# sume
Summarize emails with AI
# AI Email Summarizer

A Python tool that scans email files from date-based folders, filters them by keywords, summarizes them with OpenAI, and exports the results to **DOCX** and **EPUB**.

## Features

- Scan email files from the latest `n` days
- Filter emails by:
  - include keywords
  - exclude keywords
- Read email content from common file types such as:
  - `.txt`
  - `.eml`
  - `.html`
  - other text-based formats
- Extract a readable title from the filename
- Summarize each email using OpenAI
- Export summaries to:
  - **Word (.docx)**
  - **EPUB (.epub)**

## How It Works

The script expects your email files to be organized in folders named by date:

```text
D:\gmail_emails\
├── 20260511\
├── 20260512\
└── 20260513\
```

Each folder contains email files. The program:

1. Looks for folders in the last `n` days
2. Collects all files inside those folders
3. Filters files using include/exclude keywords
4. Reads email content
5. Sends the content to OpenAI for summarization
6. Generates a DOCX file
7. Converts the DOCX to EPUB

## Main Function

Use the main function:

```python
AI_sume(n, include_keywords, exclude_keywords)
```

### Parameters

- `n`  
  Number of recent days to scan. Default is `1` if omitted.

- `include_keywords`  
  Only emails whose filenames contain at least one of these keywords will be included.

- `exclude_keywords`  
  Emails whose filenames contain any of these keywords will be excluded.

### Examples

```python
AI_sume()
AI_sume(2)
AI_sume(2, ("bloomberg", "burry"))
AI_sume(3, None, ("Google",))
AI_sume(1, "bloomberg", "google")
```

## Configuration

Before running the script, update these paths:

```python
api_key_path = r"D:\Code\API KEY\Chatgpt API Key.txt"
EMAIL_BASE_DIR = r"D:\gmail_emails"
OUTPUT_BASE_DIR = r"D:\Articles AI-Sum\Summaries"
```

### API Key

The script reads your OpenAI API key from:

```text
D:\Code\API KEY\Chatgpt API Key.txt
```

Make sure the file exists and contains only your API key.

## Output

Generated files will be saved in a folder like:

```text
D:\Articles AI-Sum\Summaries\2026-05-13 - Summaries\
```

Example output files:

```text
2026-05-13_n5_exc-Google-burry-bloomberg_summaries.docx
2026-05-13_n5_exc-Google-burry-bloomberg_summaries.epub
```

## Requirements

Install the required Python packages:

```bash
pip install requests openai python-docx ebooklib
```

## Notes

- Emails with fewer than 50 words are skipped
- The summary is generated in Vietnamese
- The DOCX file uses each email title as a heading
- The EPUB file is built from the DOCX structure
- The script is intended for local use with folders already exported from Gmail

## Project Structure

```text
project/
├── main.py
├── README.md
└── requirements.txt
```

## Copyright

```text
Copyright (c) 2026 Your Name

All rights reserved.

This software is provided for personal and internal use only.
Unauthorized copying, modification, distribution, or commercial use
without prior written permission is prohibited.

Third-party libraries used in this project remain the property of
their respective authors and are subject to their own licenses.
```
MIT License


