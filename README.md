# File Token Counter 🔢

**File Token Counter** is a Python-based tool for counting tokens in various document formats, such as PDFs, Word files, Excel spreadsheets, and PowerPoint presentations. This application is particularly useful for those working with large language models (LLMs), providing quick insights into token usage across different document types.

## Features
- **Multi-Format Support**: Easily count tokens in `.pdf`, `.docx`, `.xlsx`, and `.pptx` files.
- **User-Friendly Interface**: Simple, responsive GUI with dynamic feedback to indicate processing status.
- **Built for LLMs**: Optimized for token models like OpenAI’s `gpt-3.5-turbo` and `gpt-4`.

## Usage
### Option 1: 
You can download the latest .exe file from the Releases section on GitHub. This option requires no Python installation or dependencies. Simply download the executable, run it, and start counting tokens in your documents.

### Option 2: 
Run the Python Script
Clone the repository:

```Bash
git clone https://github.com/Gullerboi/FileTokenCounter.git
cd FileTokenCounter
```

2. Run the program: Run the Python script directly:

```Bash
python FileTokenCounter.py
```

The program will prompt you to select a file. Once a file is selected, it displays the token count.
The tool supports PDF, Word, Excel, and PowerPoint files.
## Building the Executable
To build the .exe file yourself:

1. Install PyInstaller:
```Bash
pip install pyinstaller
```

Run the following command in the project directory:

```Bash
pyinstaller --onefile --windowed --icon=Token.ico FileTokenCounter.py
```
The .exe will be located in the dist folder. You can distribute this .exe as a standalone file.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.

## Requirements
- Python 3.7 or higher (for users running the Python script directly)
- Required packages listed in `requirements.txt`

### Dependencies
Install the dependencies by running:

```bash
pip install -r requirements.txt
```




