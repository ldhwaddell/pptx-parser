import argparse
import glob
import logging
import os
import time
from io import BytesIO
from tempfile import TemporaryDirectory
from textwrap import fill
from typing import List
from zipfile import BadZipFile, ZipExtFile, ZipFile

from pydub import AudioSegment
from rich.console import Console
from rich.logging import RichHandler
from rich.progress import Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
from rich.table import Table

import torch
from transformers import pipeline


# Set up logger
logger = logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)-8s %(message)s",
    handlers=[RichHandler()],
)

# Set up arg parser
parser = argparse.ArgumentParser(
    description="Automatic transcriptions of PowerPoint presentation embedded audio"
)

# Set up console for pretty printing
console = Console(log_time=False)


# Build speech to text pipe
pipe = pipeline(
    "automatic-speech-recognition",
    model="openai/whisper-large-v3",
    torch_dtype=torch.float16,
    device="mps",
    model_kwargs={"attn_implementation": "sdpa"},
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


def get_pptx_files(dir: str) -> List[str]:
    """
    Finds and prints all PowerPoint (.pptx) files in the given directory.

    Args:
    directory (str): The directory to search for PowerPoint files.

    Returns:
    List[str]: A list of paths to the PowerPoint files found.

    Raises:
    NotADirectoryError: If the provided path is not a directory.
    FileNotFoundError: If no .pptx files are found in the directory.
    """
    if not os.path.isdir(dir):
        raise NotADirectoryError(f"The provided path is not a directory: {dir}")

    pptx_files = glob.glob(os.path.join(dir, "*.pptx"))

    if not pptx_files:
        raise FileNotFoundError(f"No .pptx files found in {dir}")

    # Build table outline
    table = Table()
    table.add_column("Name", justify="left", style="cyan", no_wrap=True)
    table.add_column("Size", style="magenta")
    table.add_column("Created", justify="left", style="green")

    # Populate table
    for f in pptx_files:
        name = os.path.splitext(os.path.basename(f))[0]
        size = f"{os.path.getsize(f) / (1024 * 1024):.2f} mb"
        created = time.ctime(os.path.getctime(f))
        table.add_row(name, size, created)

    console.log("PowerPoints Found: ", table)

    return pptx_files


def to_wav(file: ZipExtFile, dir: str) -> str:
    """
    Converts an audio file from a zip archive to WAV format.

    Args:
    file (ZipExtFile): The audio file from the zip archive.
    dir (str): The directory to save the converted .wav file.

    Returns:
    str: The path to the converted .wav file.

    Raises:
    FileNotFoundError: If the directory to save to does not exist.
    Any other exceptions related to audio file handling.
    """
    if not os.path.isdir(dir):
        raise FileNotFoundError(f"Directory not found: {dir}")

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


def transcribe_pptx(path: str):
    # Ensure path to pptx is valid
    if not os.path.isfile(path):
        raise FileNotFoundError(f"The file {path} does not exist.")

    try:
        with ZipFile(path, "r") as zip:
            audio_files = get_audio_files_from_zip(zip)

            if not audio_files:
                logging.warning("No audio files found in the PowerPoint file.")
                return None

            # Extract audio files to a temporary directory
            with TemporaryDirectory() as temp_dir:
                transcription = ""
                for i, file in enumerate(audio_files, start=1):
                    with zip.open(file) as audio_file:
                        # Convert non .wav files to .wav
                        if not file.endswith(".wav"):
                            logging.info(f"Converting {file} to .wav")
                            wav_file_path = to_wav(audio_file, temp_dir)
                        else:
                            wav_file_path = zip.extract(file, temp_dir)

                        # Skip if error converting
                        if not wav_file_path:
                            continue

                        # Build formatted transcription
                        identifier = f"Slide {i}: "
                        text = transcribe(wav_file_path)
                        newline = "\n\n"

                        transcription += identifier + fill(text, width=100) + newline

            return transcription

    except BadZipFile:
        raise
    except Exception:
        raise


def transcribe(file_path: str) -> str:
    base_name = os.path.basename(file_path)

    start_time = time.time()

    try:
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            TimeElapsedColumn(),
            console=console,
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
    except Exception as e:
        raise Exception(f"An error occurred during transcription: {e}")


def save_transcription(dir, pptx_name, transcription):
    if not transcription:
        logging.warning(f"No transcription found for '{pptx_name}'. Skipping.")
        return

    try:
        # Ensure the directory exists
        os.makedirs(dir, exist_ok=True)
        file_path = os.path.join(dir, f"{pptx_name}.txt")

        # Write the transcription to the file
        with open(file_path, "w", encoding="utf8") as file:
            file.write(transcription)

        logging.info(f"Saved transcription of '{pptx_name}' to '{file_path}'")

    except OSError as e:
        logging.error(
            f"Failed to create/access directory '{dir}' or write to file '{file_path}': {e}"
        )
    except Exception as e:
        logging.error(f"An error occurred: {e}")


def main():
    args = parser.parse_args()
    try:
        pptx_files = get_pptx_files(args.dir)

        for f in pptx_files:
            # Get name of ppt presentation
            pptx_name = os.path.splitext(os.path.basename(f))[0]
            logging.info(f"Processing {pptx_name}...")

            # Transcribe it
            transcription = transcribe_pptx(f)

            # Save the combined text to a folder in the output directory named after the ppt presentation
            save_transcription(args.output_dir, pptx_name, transcription)
    except Exception as e:
        logging.error(e)


if __name__ == "__main__":
    main()
