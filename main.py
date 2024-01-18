import glob
import logging
import os
import time
from argparse import ArgumentParser, BooleanOptionalAction
from io import BytesIO
from tempfile import TemporaryDirectory
from textwrap import fill
from typing import List, Optional
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
                        text = transcribe(wav_file_path, pipe)
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


def transcribe(file_path: str, pipe: Pipeline) -> str:
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
    base_name = os.path.basename(file_path)
    start_time = time.time()

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        TimeElapsedColumn(),
    ) as progress:
        progress.add_task(f"[green]Transcribing {base_name}...", total=None)

        output = pipe(
            file_path,
            chunk_length_s=30,
            batch_size=4,
            return_timestamps=False,
        )

    elapsed_time = time.time() - start_time

    logging.info(
        f"Finished transcribing '{base_name}'. Duration: {elapsed_time:.2f} seconds."
    )

    return output["text"]


def save_transcription(dir: str, pptx_name: str, transcription: str) -> None:
    """
    Saves the provided transcription text to a file.

    Args:
    dir (str): The directory where the transcription file will be saved.
    pptx_name (str): The name of the PowerPoint file, which will be used as the base name for the transcription file.
    transcription (str): The transcription text to be saved.

    Returns:
    Optional[str]: The path to the saved transcription file, or None if the transcription is empty or an error occurs.

    Raises:
    OSError: If there's an issue with creating the directory or writing to the file.
    Exception: For any other issues encountered during the file writing process.
    """
    if not transcription:
        return

    try:
        # Ensure the directory exists
        os.makedirs(dir, exist_ok=True)
        file_path = os.path.join(dir, f"{pptx_name}.txt")

        # Write the transcription to the file
        with open(file_path, "w", encoding="utf8") as file:
            file.write(transcription)

        logging.info(f"Saved transcription of '{pptx_name}' to '{file_path}'\n")

    except OSError as e:
        raise OSError(
            f"Failed to create/access directory '{dir}' or write to file '{file_path}': {e}"
        )
    except Exception as e:
        raise Exception(f"An error occurred while saving the transcription: {e}")


def main():
    args = parser.parse_args()

    try:
        pptx_files = get_pptx_files(args.dir, args.recursive)

        # Build speech to text pipe only if pptx files found
        pipe = pipeline(
            "automatic-speech-recognition",
            model="openai/whisper-large-v3",
            torch_dtype=float16,
            device="mps",
            model_kwargs={"attn_implementation": "sdpa"},
        )

        for f in pptx_files:
            # Get name of ppt presentation
            pptx_name = os.path.splitext(os.path.basename(f))[0]
            logging.info(f"Processing '{pptx_name}'...")

            # Transcribe it
            transcription = transcribe_pptx(f, pipe)

            # Save the combined text to a folder in the output directory named after the ppt presentation
            save_transcription(args.output_dir, pptx_name, transcription)
    except Exception as e:
        logging.error(e)


if __name__ == "__main__":
    main()
