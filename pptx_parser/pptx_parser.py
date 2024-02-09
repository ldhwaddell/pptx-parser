import glob
import logging
import os
import time
import xml.etree.ElementTree as ET
from io import BytesIO
from tempfile import TemporaryDirectory
from typing import List, Dict, Optional
from zipfile import ZipFile

from pydub import AudioSegment


# Set up logger
logger = logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)-8s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


class PptxParser:
    """
    A parser for extracting content from PowerPoint (PPTX) files within a specified directory.

    This class provides functionality to parse PPTX files for text extraction, supporting both
    non-recursive and recursive directory searches.

    :param dir: The directory path where PPTX files are located.
    :param recursive: A boolean indicating whether the search should be recursive, default is False.
    """

    def __init__(self, dir: str, save_images_dir: str, recursive: bool = False) -> None:
        self.dir = dir
        self.save_images_dir = save_images_dir
        self.recursive = recursive
        self.pipe = None

    def parse(self) -> List[Dict]:
        """
        Parses PPTX files in the specified directory, extracting text and other relevant content.

        Searches for PPTX files based on the instance's directory and recursion settings,
        opens each PPTX file, extracts text from the archive. Extracted content is returned

        Exceptions during parsing are caught and logged as errors.

        :return List[Dict]: The list of dicts containing powerpoint slide content
        """

        try:
            pptx_files = self.__get_pptx_files()
            all_pptx_content = []
            for f in pptx_files:
                # Get name of ppt presentation
                pptx_name = os.path.splitext(os.path.basename(f))[0]
                logging.info(f"Processing '{pptx_name}'...")

                # Open zip file, get file locations of media, images, and slide .xml files
                with ZipFile(f, "r") as zip_file:
                    content_file_paths = self.__get_content_file_paths_from_zip(
                        zip_file
                    )

                    # Extract text from all the relevant files
                    pptx_content = self.__parse_content_files(
                        zip_file, content_file_paths
                    )
                    pptx_content["name"] = pptx_name
                    all_pptx_content.append(pptx_content)

            return all_pptx_content

        except Exception as e:
            logging.error(f"Error: {e}")

    def __format_slide_name(self, rel_file: str) -> str:
        """
        Extracts the slide path from the relationship file path.

        :param rel_file (str): The relationship file path.

        :return str: The slide path.
        """
        _, tail = os.path.split(rel_file)
        name, _ = os.path.splitext(tail)

        return name

    def __get_pptx_files(self) -> List[str]:
        """
        Finds all PowerPoint (.pptx) files in the given directory and returns them.
        Prints their name, creation time, and size

        :param directory (str): The directory to search for PowerPoint files.
        :param recursive (bool): Specify if dir should recursively be searched

        :returns List[str]: A list of paths to the PowerPoint files found.

        :raises NotADirectoryError: If the provided path is not a directory.
        :raises FileNotFoundError: If no .pptx files are found in the directory.
        """
        if not os.path.isdir(self.dir):
            raise NotADirectoryError(
                f"The provided path is not a directory: {self.dir}"
            )

        search_pattern = "**/*.pptx" if self.recursive else "*.pptx"
        pptx_files = glob.glob(
            os.path.join(self.dir, search_pattern), recursive=self.recursive
        )

        if not pptx_files:
            raise FileNotFoundError(f"No .pptx files found in {self.dir}")

        self.pptx_files = pptx_files

        max_name_length = 50

        for f in pptx_files:
            # Extract the base name without extension and truncate if necessary
            name = os.path.splitext(os.path.basename(f))[0]
            if len(name) > max_name_length - 3:
                name = name[: max_name_length - 3] + "..."

            # Get the creation time and format it
            created = time.ctime(os.path.getctime(f))

            # Calculate the file size in MB and format it
            size = f"{os.path.getsize(f) / (1024 * 1024):.2f} mb"

            # Print the information in a formatted way
            print(f"{name:50} {created:30} {size:>10}")

        return pptx_files

    def __get_content_file_paths_from_zip(
        self,
        zip_file: ZipFile,
    ) -> Dict[str, Dict[str, Optional[List[str]]]]:
        """
        Extracts slide, audio, and image information from a PowerPoint ZIP file.

        :param zip_file (ZipFile): A ZipFile object representing the PowerPoint file.

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
            # Add entry in dict
            slide_path = f"ppt/slides/{self.__format_slide_name(rel_file)}"
            slide_files[slide_path] = {"audio": None, "images": []}

            # Read xml and convert to tree
            xml_content = zip_file.read(rel_file).decode("utf-8")
            root = ET.fromstring(xml_content)

            # Find audio and visual assets
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

        return slide_files

    def __parse_content_files(
        self,
        zip_file: ZipFile,
        content_files: Dict[str, Dict[str, Optional[List[str]]]],
    ) -> Dict[str, Dict[str, str]]:
        """
        Parses contents of pptx from ZIP file, extracting transcriptions for audio files and text from slides.

        :param zip_file (ZipFile):  ZipFile object for pptx
        :param content_files (Dict[str, Dict[str, Optional[List[str]]]]): Dict mapping slide paths to their content
        :param pipe (Pipeline): Pipeline for audio

        :return Dict[str, Dict[str, str]]: Dict with the text and optionally transcription audio
        """
        pptx_content = {}

        for slide_path, content in content_files.items():
            slide_name = self.__format_slide_name(slide_path)
            pptx_content[slide_name] = {}

            audio_file_path = content["audio"]
            image_file_paths = content["images"]

            # Get audio
            if audio_file_path:
                pptx_content[slide_name]["transcription"] = self.__parse_audio_file(
                    audio_file_path, zip_file
                )

            # Get images
            if image_file_paths and self.save_images_dir:
                pptx_content[slide_name]["image_paths"] = self.__save_images(
                    slide_name, image_file_paths, zip_file
                )

            # Get slide text content
            pptx_content[slide_name]["slide_text"] = self.__parse_slide(
                slide_path, zip_file
            )

        return pptx_content

    def __parse_audio_file(self, audio_path: str, zip_file: ZipFile) -> str:
        """
        Processes and transcribes an audio file from a ZIP archive.

        :param audio_path (str): The path of the audio file within the ZIP archive.
        :param zip_file (ZipFile): A ZipFile object representing the ZIP archive containing the PowerPoint file.
        :param pipe (Pipeline): A transcription pipeline used for processing the audio file.

        :return: The transcribed text from the audio file.
        """
        # Import necessary library and create pipe if not already done
        if not self.pipe:
            from torch import float16
            from transformers import pipeline

            self.pipe = pipeline(
                "automatic-speech-recognition",
                model="openai/whisper-tiny",
                torch_dtype=float16,
                device="mps",
                model_kwargs={"attn_implementation": "sdpa"},
            )

        start_time = time.time()

        with TemporaryDirectory() as temp_dir:
            _, ext = os.path.splitext(audio_path)
            valid_formats = {".mp3", ".flac", ".wav"}

            if ext.lower() not in valid_formats:
                logging.info(f"Converting {os.path.basename(audio_path)} to .wav")

                # Construct .wav filename
                local_audio_path = os.path.join(
                    temp_dir,
                    os.path.splitext(os.path.basename(audio_path))[0] + ".wav",
                )

                # Read audio file as bytes
                audio = BytesIO(zip_file.read(audio_path))

                # Convert the file to WAV format
                result = AudioSegment.from_file(audio).export(
                    local_audio_path, format="wav"
                )

                # Explicitly close
                result.close()

            # If the file is already in an accepted format, just extract it
            else:
                local_audio_path = os.path.join(temp_dir, os.path.basename(audio_path))

                with zip_file.open(audio_path) as audio_file, open(
                    local_audio_path, "wb"
                ) as out_file:
                    out_file.write(audio_file.read())

            logging.info(f"Transcribing '{os.path.basename(local_audio_path)}'...")

            output = self.pipe(
                local_audio_path,
                chunk_length_s=30,
                batch_size=4,
                return_timestamps=False,
            )

        elapsed_time = time.time() - start_time

        logging.info(f"Finished transcribing. Duration: {elapsed_time:.2f} seconds.")

        return output["text"]

    def __save_images(
        self, slide_name: str, image_file_paths: List[str], zip_file: ZipFile
    ) -> List[str]:
        pass

    def __parse_slide(self, slide_file_path: str, zip_file: ZipFile) -> str:
        """
        Parses xml from a slide, extracting text

        :param slide_file_path (str): Path to xml file in zip
        :param zip_file (ZipFile):  ZipFile object for pptx

        :return str: The text from slide
        """
        # Read xml and convert to tree
        xml_content = zip_file.read(slide_file_path).decode("utf-8")
        root = ET.fromstring(xml_content)

        # Extract all text elements
        text_elements = root.findall(
            ".//a:t",
            namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"},
        )

        # Concatenate the text from each element
        slide_text = " ".join(
            [elem.text for elem in text_elements if elem.text is not None]
        )

        return slide_text
