from docx import Document
from docxcompose.composer import Composer
import argparse
from win32com import client as wc

parser = argparse.ArgumentParser()
parser.add_argument("--files", type=str, nargs='+',
                    help="input files")
parser.add_argument("--dst", type=str,
                    help="output files")
args = parser.parse_args()


def check_doc(files: list) -> list:
    docxes = []
    word = wc.Dispatch("Word.Application")
    for file in files:
        if file.endswith(".doc"):  # 排除文件夹内的其它干扰文件，只获取".doc"后缀的word文件
            docx = f"{file}x"
            doc = word.Documents.Open(file)
            doc.SaveAs(docx, 12)
            doc.Close()
            docxes.append(docx)
        elif file.endswith(".docx"):
            docxes.append(file)
        else:
            continue
    word.Quit()
    return docxes


def combined(files: list, dst: str):
    result = Document(files[0])
    result.add_page_break()
    composer = Composer(result)
    docxes = check_doc(files=files)
    print(docxes)
    for i in range(1, len(docxes)):
        doc = Document(docxes[i])

        if i != len(docxes) - 1:
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
