from os import getenv
from pathlib import Path
from typing import List
from dotenv import load_dotenv
from docx import Document
from pandas import DataFrame, ExcelWriter


def main_program() -> None:
    DOCX_FILE_PATH = get_env_path("DOCX_FILE_PATH")
    doc = Document(DOCX_FILE_PATH)
    all_paras = doc.paragraphs

    results = []

    pressure_list = []
    temperature_list = []
    value_list = []
    S_list = []
    I_list = []

    U_list = []
    Mach_list = []
    Cp_list = []
    k_list = []
    Cp_d_list = []

    k_d_list = []
    A_list = []
    Mu_list = []
    Lt_list = []
    Lt_d_list = []

    MM_list = []
    Cp_g_list = []
    k_g_list = []
    MM_g_list = []
    R_g_list = []

    Z_list = []
    Pl_list = []
    Bm_list = []
    n_list = []
    W_list = []

    W_A_ratio_list = []
    F_ratio_list = []
    F_d_list = []
    I_yd_list = []
    B_list = []

    for para in all_paras:
        flag = 0
        temp_para_list = para_sep(para.text)
        if "T=" in para.text:
            get_first_para(
                temp_para_list,
                pressure_list,
                temperature_list,
                value_list,
                S_list,
                I_list,
            )
        elif "U=" in para.text:
            get_sec_para(temp_para_list, U_list, Mach_list, Cp_list, k_list, Cp_d_list)
        elif 'k"=' in para.text:
            get_third_para(
                temp_para_list, k_d_list, A_list, Mu_list, Lt_list, Lt_d_list
            )
        elif "MM=" in para.text:
            get_fourth_para(
                temp_para_list, MM_list, Cp_g_list, k_g_list, MM_g_list, R_g_list
            )
        elif "Bm=" in para.text:
            get_fifth_para(temp_para_list, Z_list, Pl_list, Bm_list, n_list, W_list)
            if len(para.text) == 47:
                n_list.append("-")
                W_list.append("-")
                flag = 1
        elif "W/A=" in para.text:
            if len(para.text) == 62:
                B_list.append("-")
            get_sixth_list(
                temp_para_list,
                W_A_ratio_list,
                F_ratio_list,
                F_d_list,
                I_yd_list,
                B_list,
            )

        if flag:
            W_A_ratio_list.append("-")
            F_ratio_list.append("-")
            F_d_list.append("-")
            I_yd_list.append("-")
            B_list.append("-")

    RESULT_FILE_PATH = get_env_path("RESULT_FILE_PATH")
    results = [
        [
            pressure_list[i],
            temperature_list[i],
            value_list[i],
            S_list[i],
            I_list[i],
            U_list[i],
            Mach_list[i],
            Cp_list[i],
            k_list[i],
            Cp_d_list[i],
            k_d_list[i],
            A_list[i],
            Mu_list[i],
            Lt_list[i],
            Lt_d_list[i],
            MM_list[i],
            Cp_g_list[i],
            k_g_list[i],
            MM_g_list[i],
            R_g_list[i],
            Z_list[i],
            Pl_list[i],
            Bm_list[i],
            n_list[i],
            W_list[i],
            W_A_ratio_list[i],
            F_ratio_list[i],
            F_d_list[i],
            I_yd_list[i],
            B_list[i],
        ]
        for i, _ in enumerate(pressure_list)
    ]
    columns_name = [
        "p",
        "T",
        "V",
        "S",
        "I",
        "U",
        "M",
        "Cp",
        "k",
        'Cp"',
        'k"',
        "A",
        "Mu",
        "Lt",
        'Lt"',
        "MM",
        "Cp.г",
        "k.г",
        "MM.г",
        "R.г",
        "z",
        "Пл",
        "Bm",
        "n",
        "W",
        "W/A",
        "F/F*",
        'F"',
        "Iудпп",
        "B",
    ]

    write_xlsx_file(RESULT_FILE_PATH, columns_name, results)


def get_env_path(env_name: str) -> str:
    dotenv_path = Path("../.env")
    load_dotenv(dotenv_path=dotenv_path)
    path_file = getenv(env_name)

    return path_file


def para_sep(para_text: str) -> List[str]:
    para_text.split(" ")
    new_list = [i for i in para_text if i != " "]
    return new_list


def get_first_para(
    data_list: List[str],
    pressure_list: List[str],
    temperature_list: List[str],
    value_list: List[str],
    S_list: List[str],
    I_list: List[str],
) -> None:

    index_list = [i for i, el in enumerate(data_list) if el == "="]

    for i, el in enumerate(index_list):
        if i == 0:
            temp_p = data_list[el + 1 : index_list[i + 1] - 1]
            pres_val = "".join(map(str, temp_p))
            pressure_list.append(pres_val)
        elif i == 1:
            temp_T = data_list[el + 1 : index_list[i + 1] - 1]
            T_val = "".join(map(str, temp_T))
            temperature_list.append(T_val)
        elif i == 2:
            temp_V = data_list[el + 1 : index_list[i + 1] - 1]
            V_val = "".join(map(str, temp_V))
            value_list.append(V_val)
        elif i == 3:
            temp_S = data_list[el + 1 : index_list[i + 1] - 1]
            S_val = "".join(map(str, temp_S))
            S_list.append(S_val)
        elif i == 4:
            temp_I = data_list[el + 1 :]
            I_val = "".join(map(str, temp_I))
            I_list.append(I_val)


def get_sec_para(
    data_list: List[str],
    U_list: List[str],
    Mach_list: List[str],
    Cp_list: List[str],
    k_list: List[str],
    Cp_d_list: List[str],
) -> None:

    index_list = [i for i, el in enumerate(data_list) if el == "="]

    for i, el in enumerate(index_list):
        if i == 0:
            temp_U = data_list[el + 1 : index_list[i + 1] - 1]
            U_val = "".join(map(str, temp_U))
            U_list.append(U_val)
        elif i == 1:
            temp_Mach = data_list[el + 1 : index_list[i + 1] - 2]
            Mach_val = "".join(map(str, temp_Mach))
            Mach_list.append(Mach_val)
        elif i == 2:
            temp_Cp = data_list[el + 1 : index_list[i + 1] - 1]
            Cp_val = "".join(map(str, temp_Cp))
            Cp_list.append(Cp_val)
        elif i == 3:
            temp_k = data_list[el + 1 : index_list[i + 1] - 3]
            k_val = "".join(map(str, temp_k))
            k_list.append(k_val)
        elif i == 4:
            temp_Cp_d = data_list[el + 1 :]
            Cp_d_val = "".join(map(str, temp_Cp_d))
            Cp_d_list.append(Cp_d_val)


def get_third_para(
    data_list: List[str],
    k_d_list: List[str],
    A_list: List[str],
    Mu_list: List[str],
    Lt_list: List[str],
    Lt_d_list: List[str],
) -> None:

    index_list = [i for i, el in enumerate(data_list) if el == "="]

    for i, el in enumerate(index_list):
        if i == 0:
            temp_k_d = data_list[el + 1 : index_list[i + 1] - 1]
            k_d_val = "".join(map(str, temp_k_d))
            k_d_list.append(k_d_val)
        elif i == 1:
            temp_A = data_list[el + 1 : index_list[i + 1] - 2]
            A_val = "".join(map(str, temp_A))
            A_list.append(A_val)
        elif i == 2:
            temp_Mu = data_list[el + 1 : index_list[i + 1] - 2]
            Mu_val = "".join(map(str, temp_Mu))
            Mu_list.append(Mu_val)
        elif i == 3:
            temp_Lt = data_list[el + 1 : index_list[i + 1] - 3]
            Lt_val = "".join(map(str, temp_Lt))
            Lt_list.append(Lt_val)
        elif i == 4:
            temp_Lt_d = data_list[el + 1 :]
            Lt_d_val = "".join(map(str, temp_Lt_d))
            Lt_d_list.append(Lt_d_val)


def get_fourth_para(
    data_list: List[str],
    MM_list: List[str],
    Cp_g_list: List[str],
    k_g_list: List[str],
    MM_g_list: List[str],
    R_g_list: List[str],
) -> None:

    index_list = [i for i, el in enumerate(data_list) if el == "="]

    for i, el in enumerate(index_list):
        if i == 0:
            temp_MM = data_list[el + 1 : index_list[i + 1] - 4]
            MM_val = "".join(map(str, temp_MM))
            MM_list.append(MM_val)
        elif i == 1:
            temp_Cp_g = data_list[el + 1 : index_list[i + 1] - 3]
            Cp_g_val = "".join(map(str, temp_Cp_g))
            Cp_g_list.append(Cp_g_val)
        elif i == 2:
            temp_k_g = data_list[el + 1 : index_list[i + 1] - 4]
            k_g_val = "".join(map(str, temp_k_g))
            k_g_list.append(k_g_val)
        elif i == 3:
            temp_MM_g = data_list[el + 1 : index_list[i + 1] - 3]
            MM_g_val = "".join(map(str, temp_MM_g))
            MM_g_list.append(MM_g_val)
        elif i == 4:
            temp_R_g = data_list[el + 1 :]
            R_g_val = "".join(map(str, temp_R_g))
            R_g_list.append(R_g_val)


def get_fifth_para(
    data_list: List[str],
    Z_list: List[str],
    Pl_list: List[str],
    Bm_list: List[str],
    n_list: List[str],
    W_list: List[str],
) -> None:

    index_list = [i for i, el in enumerate(data_list) if el == "="]

    for i, el in enumerate(index_list):
        if i == 0:
            temp_Z = data_list[el + 1 : index_list[i + 1] - 2]
            Z_val = "".join(map(str, temp_Z))
            Z_list.append(Z_val)
        elif i == 1:
            temp_Pl = data_list[el + 1 : index_list[i + 1] - 2]
            Pl_val = "".join(map(str, temp_Pl))
            Pl_list.append(Pl_val)
        elif i == 2:
            if len(index_list) > 3:
                temp_Bm = data_list[el + 1 : index_list[i + 1] - 1]
            else:
                temp_Bm = data_list[el + 1 :]
            Bm_val = "".join(map(str, temp_Bm))
            Bm_list.append(Bm_val)
        elif i == 3:
            temp_n = data_list[el + 1 : index_list[i + 1] - 1]
            n_val = "".join(map(str, temp_n))
            n_list.append(n_val)
        elif i == 4:
            temp_W = data_list[el + 1 :]
            W_val = "".join(map(str, temp_W))
            W_list.append(W_val)


def get_sixth_list(
    data_list: List[str],
    W_A_ratio_list: List[str],
    F_ratio_list: List[str],
    F_d_list: List[str],
    I_yd_list: List[str],
    B_list: List[str],
) -> None:

    index_list = [i for i, el in enumerate(data_list) if el == "="]

    for i, el in enumerate(index_list):
        if i == 0:
            temp_W_A_ratio = data_list[el + 1 : index_list[i + 1] - 4]
            W_A_ratio_val = "".join(map(str, temp_W_A_ratio))
            W_A_ratio_list.append(W_A_ratio_val)
        elif i == 1:
            temp_F_ratio = data_list[el + 1 : index_list[i + 1] - 2]
            F_ratio_val = "".join(map(str, temp_F_ratio))
            F_ratio_list.append(F_ratio_val)
        elif i == 2:
            F_d_Bm = data_list[el + 1 : index_list[i + 1] - 4]
            F_d_val = "".join(map(str, F_d_Bm))
            F_d_list.append(F_d_val)
        elif i == 3:
            if len(index_list) > 4:
                temp_I_yd = data_list[el + 1 : index_list[i + 1] - 1]
            else:
                temp_I_yd = data_list[el + 1 :]
            I_yd_val = "".join(map(str, temp_I_yd))
            I_yd_list.append(I_yd_val)
        elif i == 4:
            temp_B = data_list[el + 1 :]
            B_val = "".join(map(str, temp_B))
            B_list.append(B_val)


def write_xlsx_file(PATH_FILE: str, columns_name: str, result: List[List[str]]) -> None:
    df = DataFrame(result, columns=columns_name)
    writer = ExcelWriter(PATH_FILE, engine="xlsxwriter")
    df.to_excel(writer, sheet_name="1", index=False)
    writer.close()


if __name__ == "__main__":
    """
    Run the program from src dir!
    """
    main_program()
