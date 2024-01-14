import os
import io
import glob
import logging
from tempfile import TemporaryDirectory
from zipfile import ZipFile, BadZipFile

import torch
from transformers import pipeline
from pydub import AudioSegment


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
        return []

    try:
        pptx_files = glob.glob(os.path.join(dir, "*.pptx"))

        return [os.path.abspath(file) for file in pptx_files]

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return []


def audio_to_wav(file, format, temp_dir):
    with io.BytesIO(file.read()) as audio_data:
        audio = AudioSegment.from_file(audio_data, format=format)
        wav_file_name = os.path.splitext(os.path.basename(file.name))[0] + ".wav"
        wav_file_path = os.path.join(temp_dir, wav_file_name)
        audio.export(wav_file_path, format="wav")

        return wav_file_path


def extract_audio_files(path: str):
    if not os.path.isfile(path):
        print(f"The file {path} does not exist.")
        return

    try:
        with ZipFile(path, "r") as zip:
            file_list = zip.namelist()

            # Filter for audio files (e.g., mp3, wav, m4a)
            audio_files = [
                f
                for f in file_list
                if f.startswith("ppt/media/") and f.endswith((".mp3", ".wav", ".m4a"))
            ]

            if not audio_files:
                logging.info("No audio files found in the PowerPoint file.")
                return

            # Extract audio files to a temporary directory
            with TemporaryDirectory() as temp_dir:
                for i, file in enumerate(audio_files, start=1):
                    try:
                        with zip.open(file) as audio_file:
                            if file.endswith(".m4a"):
                                wav_file_path = audio_to_wav(
                                    audio_file, "m4a", temp_dir
                                )
                            else:
                                wav_file_path = zip.extract(file, temp_dir)

                            text = transcribe_audio(wav_file_path)
                            with open(f"slide_{i}.txt", "w") as f:
                                f.write(text)

                    except Exception as e:
                        logging.error(f"Error processing file {file}: {e}")

    except BadZipFile:
        logging.error("The provided file is not a valid zip file.")
    except Exception as e:
        logging.error(f"An error occurred: {e}")


def transcribe_audio(path):
    transcription = pipe(
        path,
        chunk_length_s=30,
        batch_size=4,
        return_timestamps=False,
    )

    return transcription["text"]


if __name__ == "__main__":
    dir = "./powerpoints"
    pptx_files = get_pptx_files(dir)
    # for file in pptx_files:
    extract_audio_files(pptx_files[0])
