import glob
import io
import logging
import os
import sys
import time
from tempfile import TemporaryDirectory
from textwrap import fill
from zipfile import ZipFile, BadZipFile, ZipExtFile


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

    pptx_files = glob.glob(os.path.join(dir, "*.pptx"))
    return [os.path.abspath(file) for file in pptx_files]


def to_wav(file: ZipExtFile, dir: str) -> str:
    try:
        audio_data = io.BytesIO(file.read())

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

    except Exception as e:
        logging.error(f"Failed to convert {file.name} to .wav: {e}")
        return


def transcribe_pptx(path: str):
    # Ensure path to pptx is valid
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
                logging.warning("No audio files found in the PowerPoint file.")
                return

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
        logging.error("The provided file is not a valid zip file.")
        return
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return


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

    # Show elapsed transcription time
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
            file.write(transcription)

        logging.info(f"Saved transcription of '{pptx_name}' to '{file_path}'")

    except OSError as e:
        logging.error(
            f"Failed to create/access directory '{dir}' or write to file '{file_path}': {e}"
        )
        return
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return


def main():
    dir = "./powerpoints"
    output_dir = "./transcriptions"
    pptx_files = get_pptx_files(dir)

    if not pptx_files:
        logging.error("No .pptx files found")
        sys.exit(1)

    for f in pptx_files:
        # Get name of ppt presentation
        pptx_name = os.path.splitext(os.path.basename(f))[0]
        logging.info(f"Processing {pptx_name}...")

        # Transcribe it
        transcription = transcribe_pptx(f)

        # Skip if unable to transcribe
        if not transcription:
            continue

        # Save the combined text to a folder in the output directory named after the ppt presentation
        save_transcription(pptx_name, transcription, output_dir)


if __name__ == "__main__":
    main()
