# Markdown to Word Converter

## About

The Markdown to Word Converter is a Python-based tool designed to streamline the process of converting Markdown files to Microsoft Word documents. It leverages the power of Pandoc, a universal document converter, to handle the transformation while maintaining formatting and embedding images.

Key features include:
- Batch conversion of multiple Markdown files
- Flexible image handling: automatic download or manual insertion
- Support for custom input and output directories
- Preservation of Markdown formatting in the resulting Word documents

This tool accommodates various web scraping scenarios, automatically downloading images when permitted, or allowing manual image insertion when necessary. It's particularly useful for content creators, technical writers, and developers who work with Markdown for documentation but need to deliver content in Word format.

The converter is built with flexibility in mind, allowing users to easily integrate it into their existing workflows. Whether you're working with locally saved Markdown files or those created from web page captures, this tool simplifies the conversion process, making it a valuable asset in any documentation toolkit.

## Project Structure

The project follows a standard Python package structure. The `src` directory contains the main source code, with `converter.py` being the primary script to run. The `tests` directory holds the test files. The `__init__.py` files in both `src` and `tests` directories mark them as Python packages, allowing for easy imports. Input Markdown files and associated images go in the `x-INPUT` directory. Converted Word documents are output to the `x-OUTPUT` directory.
```
markdown-to-word/
│
├── src/
│   ├── __init__.py
│   └── converter.py
│
├── tests/
│   ├── __init__.py
│   └── test_converter.py
│
├── 🟤x-INPUT/
├── 🟫x-OUTPUT/
├── .gitignore
├── README.md
└── requirements.txt
```

## Installation

1. In VSCode open a new PowerShell Terminal and move to where you want to place this repository.

2. Clone this repository: 
```PowerShell
git clone https://github.com/michellepace/markdown-to-word.git
cd markdown-to-word # Move into repostory
```

3. Create a virtual environment and activate it:
```PowerShell
python -m venv .venv # Create virtual environment ".venv"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser # (if needed)
.venv\Scripts\activate # Activate virtual environment
pip list # List installed packages (will be empty)
```
4. Install the required packages: 
```PowerShell
pip install -r requirements.txt # Install libraries specified in this file
pip list # List installed packages
```
5. Install Pandoc onto your local machine:
- Download from: https://pandoc.org/installing.html
- Follow the installation instructions for your operating system

## Usage

### Preparing Markdown Files

1. To easily save web pages as Markdown files, use the Edge Browser extension [MarkDownload - Markdown Web Clipper](https://microsoftedge.microsoft.com/addons/detail/markdownload-markdown-w/hajanaajapkhaabfcofdjgjnlgkdkknm). This tool allows you to quickly capture web pages in Markdown format, streamlining the documentation collection process when scraping is restricted.

2. Download each webpage as a Markdown file and place it in the `x-INPUT\` directory.

### Handling Images

3. Image processing depends on the website's scraping policy:

   - For websites that allow scraping:
     * Pandoc will automatically download images referenced in the Markdown files.
     * These images will be embedded directly into the Word documents during conversion.
     * No manual intervention is required.

   - For websites that do not allow scraping:
     * Manually download all images referenced in the Markdown files.
     * Place these images in the `x-INPUT\` directory alongside the Markdown files.
     * Ensure you maintain the original file names of the images.
     * During conversion, Pandoc will search for these images locally and embed them into the Word documents.
   
   Note: If an image referenced in the Markdown file cannot be downloaded automatically or is not found locally, the Word document will display a clickable hyperlink to the online image.

### Converting Your Markdown Files to Word

4. Run the conversion script:
   ```PowerShell
   python src/converter.py
   ```
   
   You can also specify custom input and output directories as command-line arguments:
   ```PowerShell
   python src/converter.py  C:\my_markdown_files\   C:\my_output_folder\
   ```

5. For more options and detailed usage information, use the `--help` flag:
   ```PowerShell
   python src/converter.py --help
   ```

### Run Tests (optional)

6. Run the tests using pytest with the following command:
   ```PowerShell
   pytest tests/test_converter.py --verbase # verbose is prettier
   ```
   This will run all the unit tests defined in the test file using the pytest framework.

---

## Working Notes

### Local Git and Online GitHub

```
Your Local Machine                    GitHub (Remote)
+------------------------+            +------------------------+
|                        |            |                        |
|  Local Repository      |            |  Remote Repository     |
|  +-----------------+   |            |  +-----------------+   |
|  |                 |   |            |  |                 |   |
|  |  master branch  |   |    push    |  |  master branch  |   |
|  |     *--*--*     |---+----------->|  |     *--*--*     |   |
|  |                 |   |            |  |                 |   |
|  +-----------------+   |            |  +-----------------+   |
|                        |            |                        |
+------------------------+            +------------------------+
         ^                                       ^
         |                                       |
      "master"                              "origin/master"
```

Example usage from Git Bash Terminal:

```Bash
MP@LAPTOP-MP MINGW64 ~/OneDrive/Documentos/git-repos/markdown-to-word-midjourney (master)
$ git add README.md # Stage changes on Master (local)
$ git commit -m "Update README.md with xyz" # Commit changes to Master
$ git push -u origin master # Push changes to Origin (GitHub)
$ git status # Check status
On branch master
Your branch is up to date with 'origin/master'.
```

For `git push -u origin master`: This sets up tracking between your local repo (master) and your Github repo (origin). Once you've run this command once, you can simply use `git push` and `git push` in the future.

---