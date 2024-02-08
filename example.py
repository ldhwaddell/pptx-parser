from pptx_parser import PptxParser

# Search directory
parser = PptxParser(dir="./powerpoints", recursive=True)
parser.parse()


# Recursively search directory
recursive_parser = PptxParser(dir="./powerpoints", recursive=False)
recursive_parser.parse()


# Output results to file
parser = PptxParser(dir="./powerpoints", recursive=True)
parser.parse()


# As a CLI
from argparse import ArgumentParser, BooleanOptionalAction

# Set up arg parser
cli = ArgumentParser(description="Parse .pptx files")

# Define CLI inputs
cli.add_argument(
    "--dir",
    required=True,
    type=str,
    help="Directory where the PowerPoint files are located.",
)
cli.add_argument(
    "--output-dir",
    required=False,
    type=str,
    help="Directory where the transcriptions will be saved.",
)

cli.add_argument(
    "--recursive",
    required=False,
    action=BooleanOptionalAction,
    help="Specify if you want to recursively search dir",
)

parser = PptxParser(dir=cli.dir, recursive=cli.recursive)
parser.parse()
