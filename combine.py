from docx import Document
from docxcompose.composer import Composer
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--files", type=str, nargs='+',
                    help="input files")
parser.add_argument("--dst", type=str,
                    help="output files")
args = parser.parse_args()


def combined(files: list, dst: str):
    result = Document(files[0])
    result.add_page_break()
    composer = Composer(result)

    for i in range(1, len(files)):
        doc = Document(files[i])

        if i != len(files) - 1:
            doc.add_page_break()

        composer.append(doc)

    composer.save(dst)


def main() -> None:
    try:
        combined(files=args.files, dst=args.dst)
        print("completely")
    except Exception as error:
        print(error)


if __name__ == "__main__":
    main()
