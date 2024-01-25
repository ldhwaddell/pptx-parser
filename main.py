import glob
import logging
import os
import time
import subprocess
import xml.etree.ElementTree as ET
from argparse import ArgumentParser, BooleanOptionalAction
from io import BytesIO
from tempfile import TemporaryDirectory
from textwrap import fill
from typing import List, Optional, Dict
from zipfile import BadZipFile, ZipExtFile, ZipFile

from pydub import AudioSegment
from rich.console import Console
from rich.logging import RichHandler
from rich.progress import Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
from rich.table import Table

from torch import float16
from transformers import pipeline, Pipeline

# Set up logger
logger = logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)-8s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[RichHandler(show_time=False)],
)

# Set up arg parser
parser = ArgumentParser(
    description="Automatic transcriptions of PowerPoint presentation embedded audio"
)

# Define CLI inputs
parser.add_argument(
    "--dir",
    required=True,
    type=str,
    help="Directory where the PowerPoint files are located.",
)
parser.add_argument(
    "--output-dir",
    required=True,
    type=str,
    help="Directory where the transcriptions will be saved.",
)

parser.add_argument(
    "--recursive",
    required=False,
    action=BooleanOptionalAction,
    help="Specify if you want to recursively search dir",
)


def get_pptx_files(dir: str, recursive: bool) -> List[str]:
    """
    Finds all PowerPoint (.pptx) files in the given directory.

    Args:
    directory (str): The directory to search for PowerPoint files.
    recursive (bool): Specify if dir should recursively be searched

    Returns:
    List[str]: A list of paths to the PowerPoint files found.

    Raises:
    NotADirectoryError: If the provided path is not a directory.
    FileNotFoundError: If no .pptx files are found in the directory.
    """
    if not os.path.isdir(dir):
        raise NotADirectoryError(f"The provided path is not a directory: {dir}")

    search_pattern = "**/*.pptx" if recursive else "*.pptx"
    pptx_files = glob.glob(os.path.join(dir, search_pattern), recursive=recursive)

    if not pptx_files:
        raise FileNotFoundError(f"No .pptx files found in {dir}")

    display_table(pptx_files)

    return pptx_files


def display_table(files: List[str]) -> None:
    """
    Displays a table with the name, creation time, and size of each file in the provided list.

    Args:
    files (List[str]): A list of file paths for which details will be displayed.

    Returns:
    None: This function does not return anything.

    Note:
    This function assumes that the file paths provided in the list are valid and accessible.
    It does not handle exceptions related to invalid file paths or inaccessible files.
    """
    # Set up console for pretty printing
    console = Console(log_time=False)

    # Build table outline
    table = Table()
    table.add_column("Name", justify="left", style="cyan", no_wrap=True)
    table.add_column("Created", justify="left", style="green")
    table.add_column("Size", style="magenta")

    # Populate table
    for f in files:
        name = os.path.splitext(os.path.basename(f))[0]
        created = time.ctime(os.path.getctime(f))
        size = f"{os.path.getsize(f) / (1024 * 1024):.2f} mb"
        table.add_row(name, created, size)

    console.log("PowerPoints Found: ", table)


def format_slide_name(rel_file: str) -> str:
    """
    Extracts the slide path from the relationship file path.

    :param rel_file: The relationship file path.
    :return: The slide path.
    """
    _, tail = os.path.split(rel_file)
    name, _ = os.path.splitext(tail)
    return name


def process_relationships(
    zip_file: ZipFile, rel_file: str, slide_files: Dict, slide_path: str
) -> None:
    """
    Processes relationships within a PowerPoint relationship file.

    :param zip_file: The ZipFile object.
    :param rel_file: The relationship file path.
    :param slide_files: The dictionary to store slide information.
    :param slide_path: The path of the slide.
    """

    xml_content = zip_file.read(rel_file).decode("utf-8")
    root = ET.fromstring(xml_content)

    for rel in root.findall(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    ):
        rel_type = rel.get("Type")

        if "relationships/audio" in rel_type:
            target = rel.get("Target").replace("../", "ppt/")
            slide_files[slide_path]["audio"] = target

        elif "relationships/image" in rel_type:
            target = rel.get("Target").replace("../", "ppt/")
            slide_files[slide_path]["images"].append(target)


def get_content_file_paths_from_zip(
    zip_file: ZipFile,
) -> Dict[str, Dict[str, Optional[List[str]]]]:
    """
    Extracts slide, audio, and image information from a PowerPoint ZIP file.

    :param zip_file: A ZipFile object representing the PowerPoint file.
    :return: A dictionary mapping slide paths to their corresponding audio and image files.
    :raises NotADirectoryError: If the required directory is not found in the zip file.
    """

    rels_dir = "ppt/slides/_rels"

    files = zip_file.namelist()

    if not any(item.startswith(rels_dir) for item in files):
        raise NotADirectoryError(
            f"Unable to find '{rels_dir}' directory in {zip_file.filename}"
        )

    slide_files = {}
    for rel_file in (f for f in files if f.startswith(rels_dir)):
        slide_path = f"ppt/slides/{format_slide_name(rel_file)}"
        slide_files[slide_path] = {"audio": None, "images": []}
        process_relationships(zip_file, rel_file, slide_files, slide_path)

    return slide_files


def to_wav(file: ZipExtFile, dir: str) -> str:
    """
    Converts an audio file from a zip archive to WAV format.

    Args:
    file (ZipExtFile): The audio file from the zip archive.
    dir (str): The directory to save the converted .wav file.

    Returns:
    str: The path to the converted .wav file.

    Raises:
    Any exceptions related to audio file handling.
    """
    # Let user know conversion is happening
    logging.info(f"Converting {file.name} to .wav")

    audio_data = BytesIO(file.read())

    # Get format from file extension
    file_format = os.path.splitext(file.name)[1].strip(".")
    audio: AudioSegment = AudioSegment.from_file(audio_data, format=file_format)

    # Construct .wav filename
    wav_file_name = os.path.splitext(os.path.basename(file.name))[0] + ".wav"
    wav_file_path = os.path.join(dir, wav_file_name)

    # Export audio as .wav file
    audio.export(wav_file_path, format="wav")
    logging.info(f"Converted {file.name} to {wav_file_path}")

    return wav_file_path


def get_audio_files_from_zip(zip: ZipFile) -> List[str]:
    """
    Extracts a list of audio file paths from a ZipFile object.

    Args:
    zip (ZipFile): A ZipFile object opened in read mode.

    Returns:
    List[str]: A sorted list of audio file paths found in the ZipFile.
    """
    file_list = zip.namelist()

    # Filter for audio files (e.g., mp3, wav, m4a)
    # Sort audio files based on file name - ex: 'media2.m4a', 'media1.m4a' -> 'media1.m4a', 'media2.m4a'
    audio_files = sorted(
        [
            f
            for f in file_list
            if f.startswith("ppt/media/") and f.endswith((".mp3", ".wav", ".m4a"))
        ]
    )

    return audio_files


def transcribe_pptx(path: str, pipe: Pipeline) -> Optional[str]:
    """
    Transcribes audio files embedded in a PowerPoint (.pptx) file.

    Args:
    path (str): Path to the PowerPoint file.
    pipe (Pipeline): The pipeline to use whisper

    Returns:
    Optional[str]: The combined transcription of all audio files, or None if no audio files are found.

    Raises:
    BadZipFile: If the PowerPoint file is not a valid zip file.
    Exception: For any other issues encountered during processing.
    """
    try:
        # Open zip file, extract audio files
        zip = ZipFile(path, "r")
        audio_files = get_audio_files_from_zip(zip)

        if not audio_files:
            logging.warning("No audio files found in the PowerPoint file.")
            return None

        # Extract audio files to a temporary directory
        with TemporaryDirectory() as temp_dir:
            transcription = ""

            # Open audio files
            for i, f in enumerate(audio_files, start=1):
                with zip.open(f, "r") as audio_file:
                    try:
                        # Convert non .wav files to .wav in temp dir
                        if not f.endswith(".wav"):
                            wav_file_path = to_wav(audio_file, temp_dir)
                        else:
                            wav_file_path = zip.extract(f, temp_dir)

                        # Build formatted transcription
                        identifier = f"Slide {i}: "
                        text = "transcribe(wav_file_path, pipe)"
                        newline = "\n\n"

                        transcription += identifier + fill(text, width=100) + newline

                    except Exception as e:
                        logging.error(f"Error with audio file: {e}. Skipping")
                        continue

        return transcription

    except BadZipFile:
        raise
    except Exception:
        raise
    finally:
        zip.close()


def transcribe(audio: bytes, pipe: Pipeline) -> str:
    """
    Transcribes audio content from a given file.

    Args:
    file_path (str): The path to the audio file to be transcribed.
    pipe (Pipeline): The pipeline to use whisper

    Returns:
    str: The transcribed text from the audio file

    Raises:
    Any exceptions related to transcription.
    """
    # base_name = os.path.basename(file_path)
    start_time = time.time()

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        TimeElapsedColumn(),
    ) as progress:
        progress.add_task(f"[green]Transcribing...", total=None)

        output = pipe(
            audio,
            chunk_length_s=30,
            batch_size=4,
            return_timestamps=False,
        )

    elapsed_time = time.time() - start_time

    # logging.info(
    #     f"Finished transcribing '{base_name}'. Duration: {elapsed_time:.2f} seconds."
    # )

    return output["text"]



def convert_to_wav(input_file_path, output_file_path):
    """
    Converts an audio file to WAV format using FFmpeg.

    Args:
        input_file_path (str): Path to the input audio file.
        output_file_path (str): Path where the output WAV file will be saved.
    """
    ffmpeg_command = [
        "ffmpeg",
        "-i", input_file_path,
        "-acodec", "pcm_s16le",  # PCM 16-bit little-endian codec
        "-ar", "44100",  # Sampling rate
        "-ac", "1",  # Number of audio channels
        output_file_path
    ]

    subprocess.run(ffmpeg_command, check=True)


def parse_content_files(
    zip_file: ZipFile,
    content_files: Dict[str, Dict[str, Optional[List[str]]]],
    pipe: Pipeline,
) -> str:
    for slide_path, content in content_files.items():
        slide_name = format_slide_name(slide_path)

        audio_file_path = content["audio"]

        if audio_file_path:
            # audio = zip_file.read(audio_file_path)


            with TemporaryDirectory() as temp_dir:
                _, ext = os.path.splitext(audio_file_path)

                if ext.lower() not in [".mp3", ".flac", ".wav"]:
                    # If the file is not in one of the accepted formats, convert it to WAV
                    local_audio_path = os.path.join(temp_dir, os.path.basename(audio_file_path) + ".wav")
                    with zip_file.open(audio_file_path) as audio_file:
                        with open(local_audio_path, "wb") as out_file:
                            out_file.write(audio_file.read())

                    # Convert the file to WAV format
                    convert_to_wav(local_audio_path, local_audio_path)
                else:
                    # If the file is already in an accepted format, just extract it
                    local_audio_path = os.path.join(temp_dir, os.path.basename(audio_file_path))
                    with zip_file.open(audio_file_path) as audio_file, open(local_audio_path, "wb") as out_file:
                        out_file.write(audio_file.read())

                output = pipe(
                    local_audio_path,
                    chunk_length_s=30,
                    batch_size=4,
                    return_timestamps=False,
                )

                print(output)



        break


def main():
    args = parser.parse_args()

    try:
        pptx_files = get_pptx_files(args.dir, args.recursive)

        # Build speech to text pipe only if pptx files found
        pipe = pipeline(
            "automatic-speech-recognition",
            model="openai/whisper-tiny",
            torch_dtype=float16,
            device="mps",
            model_kwargs={"attn_implementation": "sdpa"},
        )

        for f in pptx_files:

            # Get name of ppt presentation
            pptx_name = os.path.splitext(os.path.basename(f))[0]
            logging.info(f"Processing '{pptx_name}'...")

            # Open zip file, get file locations of media, images, and slide'x'.xml files
            zip_file = ZipFile(f, "r")
            content_file_paths = get_content_file_paths_from_zip(zip_file)

            # Extract text from all the relevant files
            slides_text = parse_content_files(zip_file, content_file_paths, pipe)
            print(slides_text)

    except Exception as e:
        logging.error(f"Error: {e}")


if __name__ == "__main__":
    main()
