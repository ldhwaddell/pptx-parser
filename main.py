import glob
import io
import logging
import os
import sys
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
# pipe = pipeline(
#     "automatic-speech-recognition",
#     model="openai/whisper-large-v3",
#     torch_dtype=torch.float16,
#     device="mps",
#     model_kwargs={"attn_implementation": "sdpa"},
# )


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


def transcribe_audio_files(path: str):
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
                transcriptions = []
                for file in audio_files:
                    try:
                        with zip.open(file) as audio_file:
                            if file.endswith(".m4a"):
                                wav_file_path = audio_to_wav(
                                    audio_file, "m4a", temp_dir
                                )
                            else:
                                wav_file_path = zip.extract(file, temp_dir)

                            # text = transcribe(wav_file_path)
                            transcription = {
                                "audio_path": wav_file_path,
                                "text": "text",
                            }
                            transcriptions.append(transcription)

                    except Exception as e:
                        logging.error(f"Error processing file {file}: {e}")

            return transcriptions

    except BadZipFile:
        logging.error("The provided file is not a valid zip file.")
    except Exception as e:
        logging.error(f"An error occurred: {e}")


def transcribe(path):
    output = pipe(
        path,
        chunk_length_s=30,
        batch_size=4,
        return_timestamps=False,
    )

    return output["text"]


def save_transcriptions(transcriptions, dir):
    if not transcriptions:
        logging.info("No transcriptions found")
        return

    if not os.path.exists(dir):
        try:
            os.makedirs(dir)
            logging.info(f"Created directory: {dir}")
        except OSError as e:
            logging.error(f"Failed to create directory {dir}: {e}")
            return
    elif not os.path.isdir(dir):
        logging.error(f"The path {dir} is not a directory")
        return

    for transcription in transcriptions:
        id = transcription["path"].split(".")[0][-1]
        file_path = os.path.join(dir, f"slide_{id}.txt")

        try:
            with open(file_path, "w") as file:
                file.write(str(transcription))
            logging.info(f"Saved transcription to {file_path}")
        except IOError as e:
            logging.error(f"Failed to write file {file_path}: {e}")


if __name__ == "__main__":
    dir = "./powerpoints"
    output_dir = "./transcriptions"
    pptx_files = get_pptx_files(dir)
    if not pptx_files:
        logging.error("No .pptx files found")
        sys.exit(1)

    for f in pptx_files:
        pptx_name = os.path.basename(f).split(".")[0]
        print(pptx_name)
        # transcriptions = transcribe_audio_files(pptx_files[0])
        # # print(transcriptions)

        # save_transcriptions(transcriptions, dir=output_dir)
        break

""" 
Instead of having a .txt file for each slide. Instead, combine them into one text file that has the same name as the original ppt. 
separate the slides just with spacing and put "slide x" before the text of each. 

"""
