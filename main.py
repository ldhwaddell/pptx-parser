import glob
import io
import logging
import os
import sys
import time
from tempfile import TemporaryDirectory
from textwrap import fill
from zipfile import ZipFile, BadZipFile


from pydub import AudioSegment
from rich.progress import (
    Progress,
    TimeElapsedColumn,
    TextColumn,
    SpinnerColumn,
)
import torch
from transformers import pipeline


# Set up logger
logger = logging.basicConfig(
    level=logging.INFO, format="%(asctime)s %(name)-8s %(levelname)-8s %(message)s"
)

# Build speech to text pipe
pipe = pipeline(
    "automatic-speech-recognition",
    model="openai/whisper-large-v3",
    torch_dtype=torch.float16,
    device="mps",
    model_kwargs={"attn_implementation": "sdpa"},
)


def get_pptx_files(dir: str):
    if not os.path.isdir(dir):
        logging.error(f"The provided path is not a directory: {dir}")
        return

    try:
        pptx_files = glob.glob(os.path.join(dir, "*.pptx"))
        return [os.path.abspath(file) for file in pptx_files]

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return


def audio_to_wav(file, format, dir):
    with io.BytesIO(file.read()) as audio_data:
        audio = AudioSegment.from_file(audio_data, format=format)
        wav_file_name = os.path.splitext(os.path.basename(file.name))[0] + ".wav"
        wav_file_path = os.path.join(dir, wav_file_name)
        audio.export(wav_file_path, format="wav")

        return wav_file_path


def transcribe_audio_files(path: str):
    if not os.path.isfile(path):
        logging.error(f"The file {path} does not exist.")
        return

    try:
        with ZipFile(path, "r") as zip:
            file_list = zip.namelist()

            # Filter for audio files (e.g., mp3, wav, m4a)
            # Sort audio files based on file name - ex: 'media1.m4a', 'media3.m4a', 'media2.m4a' -> 'media1.m4a', 'media2.m4a', 'media3.m4a'
            audio_files = sorted(
                [
                    f
                    for f in file_list
                    if f.startswith("ppt/media/")
                    and f.endswith((".mp3", ".wav", ".m4a"))
                ]
            )

            if not audio_files:
                logging.info("No audio files found in the PowerPoint file.")
                return

            # Extract audio files to a temporary directory
            with TemporaryDirectory() as temp_dir:
                transcription = ""
                for i, file in enumerate(audio_files, start=1):
                    with zip.open(file) as audio_file:
                        if file.endswith(".m4a"):
                            logging.info(f"Converting {file} to .wav")
                            wav_file_path = audio_to_wav(audio_file, "m4a", temp_dir)
                        else:
                            wav_file_path = zip.extract(file, temp_dir)

                        identifier = f"Slide {i}: "
                        text = transcribe(wav_file_path)
                        newline = "\n\n"

                        transcription += identifier + text + newline

            return transcription

    except BadZipFile:
        logging.error("The provided file is not a valid zip file.")
    except Exception as e:
        logging.error(f"An error occurred: {e}")


def transcribe(file_path):
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

    # Inform the user that the transcription has finished and show the elapsed time
    logging.info(
        f"Finished transcribing '{base_name}'. Duration: {elapsed_time:.2f} seconds."
    )

    return output["text"]


def save_transcription(pptx_name, transcription, dir):
    if not transcription:
        logging.info(f"No transcription found for '{pptx_name}'")
        return

    try:
        # Ensure the directory exists
        os.makedirs(dir, exist_ok=True)
        file_path = os.path.join(dir, f"{pptx_name}.txt")

        # Write the transcription to the file
        with open(file_path, "w", encoding="utf8") as file:
            # TODO: Stop fill from removing the newlines to separate slides
            file.write(fill(transcription, width=100))

        logging.info(f"Saved transcription of '{pptx_name}' to '{file_path}'")

    except OSError as e:
        logging.error(
            f"Failed to create/access directory '{dir}' or write to file '{file_path}': {e}"
        )


if __name__ == "__main__":
    dir = "./powerpoints"
    output_dir = "./transcriptions"
    pptx_files = get_pptx_files(dir)

    if not pptx_files:
        logging.error("No .pptx files found")
        sys.exit(1)

    for f in pptx_files:
        # Get name of the actual ppt presentation
        pptx_name = os.path.basename(f).split(".")[0]
        logging.info(f"Transcribing {pptx_name}...")

        # Transcribe it
        transcription = transcribe_audio_files(f)

        # Save the combined text to a folder in the output directory named after the ppt presentation
        save_transcription(pptx_name, transcription, output_dir)
