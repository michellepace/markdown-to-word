import argparse
import textwrap
from pathlib import Path
import subprocess
import tempfile
import re


def convert_markdown_to_word(markdown_file: Path, output_file: Path):
    """
    Convert a Markdown file to a Word document using Pandoc.

    :param markdown_file: Path to the input Markdown file
    :param output_file: Path to the output Word document
    :raises subprocess.CalledProcessError: If Pandoc conversion fails
    :raises PermissionError: If the output file is already open
    """
    # Check if the file is already open
    if output_file.exists():
        try:
            output_file.rename(output_file)
        except PermissionError:
            raise PermissionError(f"Unable to write to \"{output_file}\" - you probably have it open!")

    with tempfile.TemporaryDirectory() as temp_dir: # for images pandoc might download
        cmd = [
            'pandoc', # The command to run Pandoc
            str(markdown_file), # Input file
            '-o', str(output_file), # Output file
            f'--extract-media={temp_dir}',  # Automatic image download (if permitted) into temp directory.
            '--resource-path', str(markdown_file.parent), # Tell Pandoc where to find local images
            '--wrap=preserve', # Preserve line breaks from original Markdown in the output
            '-t', 'docx'  # Make output format as Word document
        ]

        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True) # run pandoc
        except subprocess.CalledProcessError as e:
            print(f"Error converting {markdown_file}: {e}")
            print(f"Pandoc output: {e.output}")
            raise


def _find_local_image(file_name: str, markdown_file_path: Path) -> Path | None:
    """
    Find a local file with a case-insensitive match to the image filename.
    Searches in the same directory as the Markdown file.
    """
    return next((f for f in markdown_file_path.parent.iterdir() if f.name.lower() == file_name.lower()), None)


def correct_image_paths(markdown_content: str, markdown_file_path: Path) -> str:
    """
    Correct image paths in Markdown content to use local files when available.
    If a local file is not found, convert the image syntax to a downloadable link.
    
    :param markdown_content: Original Markdown content
    :param markdown_file_path: Path to the Markdown file (images are expected to be in the same directory)
    :return: Markdown content with corrected local image paths or downloadable links
    """
    def replace_path(match):
        alt_text, old_path = match.groups()
        file_name = Path(old_path).name
        local_file = _find_local_image(file_name, markdown_file_path)
        # Link with local file target else original url (text) only
        return f'![{alt_text}]({local_file})' if local_file else f"[CLICK TO VIEW ONLINE IMAGE: ({alt_text})]({old_path})"
        
    pattern_markdown_image = r'!\[(.*?)\]\((https?://.*?)\)'
    return re.sub(pattern_markdown_image, replace_path, markdown_content)


def _process_single_markdown_file(markdown_file: Path, output_dir: Path) -> None:
    """
    Process a single Markdown file: correct image paths and convert to Word.
    Assumes images are in the same directory as the Markdown file.

    :param markdown_file: Path to the input Markdown file
    :param output_dir: Directory for output Word documents
    :raises subprocess.CalledProcessError: If Word conversion fails
    :raises PermissionError: If the output file is already open
    """
    content = markdown_file.read_text(encoding='utf-8')
    content_with_local_images = correct_image_paths(content, markdown_file)

    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.md', delete=False) as temp_file:
        temp_file.write(content_with_local_images)
        temp_md_file = Path(temp_file.name)

    try:
        output_file = output_dir / (markdown_file.stem + '.docx')
        convert_markdown_to_word(temp_md_file, output_file)
        print(f"  ✅ Converted {markdown_file.name} to {output_file.name}")
    except PermissionError as e:
        print(f"  ❌ Failed to convert {markdown_file.name}: {str(e)}")
    except subprocess.CalledProcessError:
        print(f"  ❌ Failed to convert {markdown_file.name}: Pandoc error occurred")
    finally:
        temp_md_file.unlink()


def process_markdown_files(input_path: Path, output_dir: Path) -> None:
    """
    Process Markdown files from the input path.
    
    :param input_path: Path to input file or directory
    :param output_dir: Directory for output Word documents
    """
    if input_path.is_file():
        if input_path.suffix.lower() == '.md':
            _process_single_markdown_file(input_path, output_dir)
        else:
            print(f"  ❌ \"{input_path}\" is not a Markdown file.")
    elif input_path.is_dir():
        markdown_files = list(input_path.glob('*.md'))
        if not markdown_files:
            print(f"  ❌ No Markdown files found in folder '{input_path}'")
            return
        for markdown_file in markdown_files:
            try:
                _process_single_markdown_file(markdown_file, output_dir)
            except Exception as e:
                print(f"  ❌ Error processing '{markdown_file.name}': {str(e)}")
    else:
        print(f"Error: '{input_path}' is neither a file nor a directory.")


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Convert Markdown files to Word documents.",
        epilog=textwrap.dedent("""
        ----------------------------------------------------------------------------------------------
        🎯 Usage Examples:
           • python converter.py                             # Use default input and output folders
           • python converter.py input_folder output_folder  # Convert all .md files in input_folder
           • python converter.py input_file.md output_folder # Convert a single Markdown file
        ----------------------------------------------------------------------------------------------
        """),
        formatter_class=argparse.RawTextHelpFormatter
    )

    parser.add_argument("input", nargs="?", default="x-INPUT",
                        help=textwrap.dedent("""
                        Input file or directory. Can be a single .md file
                        or a folder containing .md files. If not specified,
                        the default "x-input" folder will be used.
                        """))
    parser.add_argument("output", nargs="?", default="x-OUTPUT",
                        help=textwrap.dedent("""
                        Output directory for the converted Word documents.
                        If not specified, the default "x-output" folder 
                        will be used.
                        """))
    args = parser.parse_args()

    # Ensure that either both or no arguments are provided
    if (args.input == "x-INPUT") != (args.output == "x-OUTPUT"):
        parser.error("\n\t❌ Either specify both input and output, or neither.\n\t💡 For help run 'python converter.py --help'\n")

    return args


def main() -> None:
    """
    Set up directories and initiate Markdown processing based on command-line arguments.
    """
    args = parse_arguments()
    
    input_path = Path(args.input)
    output_dir = Path(args.output)

    # Only create input directory if it's the default and doesn't exist
    if args.input == "x-INPUT" and not input_path.exists():
        input_path.mkdir(exist_ok=True)

    if not input_path.exists():
        print(f" ❌ Input path '{input_path}' does not exist.")
        print(f" 💡 For help run 'python converter.py --help'\n")
        return

    # Create output directory if it doesn't exist
    output_dir.mkdir(exist_ok=True)

    print(f"\n🟤 Processing input from:\n   {input_path.resolve()}")
    process_markdown_files(input_path, output_dir)
    print(f"\n🟫 Output directory for converted docx:\n   {output_dir.resolve()}")
    print()



if __name__ == "__main__":
    main()
