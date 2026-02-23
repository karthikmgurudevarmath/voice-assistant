# Voice Assistant Setup Guide

This guide helps you set up and run the python-based Voice Assistant extracted from your documentation.

## Prerequisites

- Python 3.x installed (Make sure to add Python to PATH during installation)
- An active internet connection
- A microphone

## Installation

1.  **Open Command Prompt/PowerShell** in the `D:\voiceaa` directory.
2.  **Install Dependencies**:
    Run the following command to install all required libraries:
    ```bash
    pip install -r requirements.txt
    ```

    *Note: If you encounter errors installing `pyaudio`, you may need to install the pre-compiled binary manually directly from [lfd.uci.edu/~gohlke/pythonlibs/](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio) matching your Python version/architecture.*

## Configuration

The script uses several APIs that require keys. Open `voice_assistant.py` in a text editor (like VS Code or Notepad) and find the following placeholders to update:

-   **WolframAlpha**: Search for `app_id = "Wolframalpha api id"` and replace with your App ID.
-   **OpenWeatherMap**: Search for `api_key = "Api key"` (inside the weather section) and replace with your key.
-   **Twilio** (Optional, for SMS): Update `account_sid`, `auth_token`, `from_`, and `to` fields in the "send message" section.
-   **Gmail** (Optional, for Email): Update `server.login('your_email_id', 'your_email_password')` in the `sendEmail` function. *Note: You may need to generate an "App Password" in your Google Account security settings for this to work.*
-   **NewsAPI**: Search for `apiKey=YOUR_NEWS_API_KEY` in the news section and replace it.

## Running the Assistant

Execute the script:
```bash
python voice_assistant.py
```

## Troubleshooting

-   **"Unable to Recognize"**: Ensure your microphone is working and background noise is minimal.
-   **Import Errors**: If it says "Module not found", ensure `pip install -r requirements.txt` completed successfully.
