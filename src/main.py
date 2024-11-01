from os import getenv
from pathlib import Path
from typing import List
from dotenv import load_dotenv
from docx import Document
from pandas import DataFrame, ExcelWriter


def main_program():
    DOCX_FILE_PATH = get_env_path("DOCX_FILE_PATH")

    doc = Document(DOCX_FILE_PATH)

    all_paras = doc.paragraphs
    len(all_paras)

    for para in all_paras:
        print(para.text[1:10])
        print("----------------")

    # ch_res_res = List[float]
    # column_names = ["", "", ""]
    # write_xlsx_data(RESULT_FILE_PATH, column_names, "first_sheet", ch_res_res)


def get_env_path(env_name: str) -> str:
    dotenv_path = Path("../.env")
    load_dotenv(dotenv_path=dotenv_path)
    path_file = getenv(env_name)

    return path_file


def write_xlsx_data(
    PATH_FILE: str, column_names: str, sheet_number: str, result: List[List[int]]
) -> None:
    df = DataFrame(result, columns=column_names)
    writer = ExcelWriter(PATH_FILE, engine="xlsxwriter")
    df.to_excel(writer, sheet_name=sheet_number, index=False)
    writer.close()


if __name__ == "__main__":
    """
    Run the program from src dir!
    """
    main_program()
