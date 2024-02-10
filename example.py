from pptx_parser import PptxParser

# Search directory
parser = PptxParser(dir="./powerpoints")
parser.parse()


# Recursively search directory
recursive_parser = PptxParser(dir="./powerpoints", recursive=True)
recursive_parser.parse()

# Transcribe audio
audio_parser = PptxParser(dir="./powerpoints", transcribe_audio=True, recursive=True)
audio_parser.parse()

# Save images
image_parser = PptxParser(
    dir="./powerpoints", save_images_dir="./output", recursive=True
)
audio_parser.parse()
