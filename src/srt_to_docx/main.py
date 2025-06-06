import asyncio
import datetime
import io
import os
import sys
import warnings

import srt

# Suppress pkg_resources deprecation warning from docxcompose
warnings.filterwarnings("ignore", category=UserWarning)
from docxtpl import DocxTemplate  # noqa: E402

DOCX_TEMPLATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "__assets__", "srt-to.docx")


def remove_microseconds(td):
    """Remove microseconds from timedelta"""
    return td - datetime.timedelta(microseconds=td.microseconds)


def process_subtitles(subtitle_generator):
    """Generator that processes subtitles to remove microseconds and leading spaces"""
    for subtitle in subtitle_generator:
        # Remove microseconds from start and end times
        subtitle.start = remove_microseconds(subtitle.start)
        subtitle.end = remove_microseconds(subtitle.end)
        # Remove leading spaces from content
        subtitle.content = subtitle.content.lstrip()
        yield subtitle


async def convert_srt_to_docx(file_path, template_bytes):
    filename_without_ext = os.path.splitext(os.path.basename(file_path))[0]
    file_path_without_ext = os.path.splitext(file_path)[0]

    # Read SRT file
    with open(file_path, "r", encoding="utf-8") as f:
        subtitle_generator = srt.parse(f.read())

    # Create DocxTemplate from bytes
    template_io = io.BytesIO(template_bytes)
    context = {
        "filename": filename_without_ext,
        "subtitles": process_subtitles(subtitle_generator),
    }
    doc = DocxTemplate(template_io)
    doc.render(context)

    try:
        doc.save(file_path_without_ext + ".docx")
        print(f"✓ Converted: {filename_without_ext}.srt -> {filename_without_ext}.docx")
    except PermissionError:
        print(
            f"✗ Permission denied: Cannot save '{filename_without_ext}.docx'. Please close the file if it's open and try again."
        )
        return False

    return True


async def convert_single(path):
    with open(DOCX_TEMPLATE_FILE, "rb") as f:
        template_bytes = f.read()
    result = await convert_srt_to_docx(path, template_bytes)
    if not result:
        sys.exit(1)


async def find_and_convert_srt_files(directory):
    # Read template once into memory
    with open(DOCX_TEMPLATE_FILE, "rb") as f:
        template_bytes = f.read()

    # Collect all SRT file paths
    srt_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".srt"):
                srt_files.append(os.path.join(root, file))

    if not srt_files:
        print("No .srt files found in the directory.")
        return

    # Process all files concurrently
    tasks = [convert_srt_to_docx(file_path, template_bytes) for file_path in srt_files]
    results = await asyncio.gather(*tasks)

    # Check if any conversions failed
    failed_count = sum(1 for result in results if not result)
    success_count = len(results) - failed_count

    print(f"\nCompleted: {success_count} successful, {failed_count} failed")
    if failed_count > 0:
        print("Some files could not be converted due to permission errors.")


def main():
    if len(sys.argv) == 1:
        asyncio.run(find_and_convert_srt_files(os.getcwd()))
    elif len(sys.argv) != 2:
        print("Usage: srt-to-docx [directory_path|file_path]")
    else:
        path = sys.argv[1]
        if not os.path.exists(path):
            print(f"Error: Path '{path}' does not exist.")
            sys.exit(1)
        elif os.path.isfile(path):
            if path.endswith(".srt"):
                asyncio.run(convert_single(path))
            else:
                print(f"Error: '{path}' is not an .srt file.")
                sys.exit(1)
        elif os.path.isdir(path):
            asyncio.run(find_and_convert_srt_files(path))
        else:
            print(f"Error: '{path}' is neither a file nor a directory.")
            sys.exit(1)


if __name__ == "__main__":
    main()
