import pandas as pd
import numpy as np


def calculate_level(col_name) -> str:
    conditions = [
        (col_name <= 1.5),
        (col_name > 1.5) & (col_name < 2.5),  # работает только с битовым &
        (col_name >= 2.5),
    ]
    results = ["Н", "С", "В"]
    return np.select(conditions, results)


def make_students_df(source_file_name):
    df = pd.read_excel(source_file_name)

    df.index += 1  # у первого ученика индекс 1
    df.index.name = "№ пп"

    criteria_1 = [
        "предметные 1",
        "предметные 2",
        "предметные 3"
    ]
    criteria_2 = [
        "метапредметные 1",
        "метапредметные 2",
        "метапредметные 3"
    ]
    criteria_3 = [
        "личностные 1",
        "личностные 2",
        "личностные 3"
    ]
    criteria_total = [
        "Предметные результаты (П)",
        "Метапредметные результаты (М)",
        "Личностные результаты (Л)"
    ]

    df["Предметные результаты (П)"] = df[criteria_1].mean(axis=1)
    df["Предметный уровень"] = calculate_level(
        df["Предметные результаты (П)"]
    )

    df["Метапредметные результаты (М)"] = df[criteria_2].mean(axis=1)
    df["Метапредметный уровень"] = calculate_level(
        df["Метапредметные результаты (М)"]
    )

    df["Личностные результаты (Л)"] = df[criteria_3].mean(axis=1)
    df["Личностный уровень"] = calculate_level(
        df["Личностные результаты (Л)"]
    )

    df["В баллах (П + М + Л) / 3"] = df[criteria_total].mean(axis=1)
    df["Уровень начальной подготовки"] = calculate_level(
        df["В баллах (П + М + Л) / 3"]
    )

    return df


def make_total_df(students_df):
    df1 = pd.DataFrame()
    df1["Предметные уровни"] = students_df["Предметный уровень"].value_counts()
    df1["Метапредметные уровни"] = (
        students_df["Метапредметный уровень"].value_counts()
    )
    df1["Личностные уровни"] = students_df["Личностный уровень"].value_counts()
    df1["Итоговые уровни"] = (
        students_df["Уровень начальной подготовки"].value_counts()
    )
    df1["Предметные уровни"] = df1["Предметные уровни"].fillna(0)
    df1["Метапредметные уровни"] = df1["Метапредметные уровни"].fillna(0)
    df1["Личностные уровни"] = df1["Личностные уровни"].fillna(0)
    df1["Итоговые уровни"] = df1["Итоговые уровни"].fillna(0)

    df1["Предметные проценты"] = (
        df1["Предметные уровни"] * 100 / len(students_df.index)
    )
    df1["Метапредметные проценты"] = (
        df1["Метапредметные уровни"] * 100 / len(students_df.index)
    )
    df1["Личностные проценты"] = (
        df1["Личностные уровни"] * 100 / len(students_df.index)
    )
    df1["Итоговые проценты"] = (
        df1["Итоговые уровни"] * 100 / len(students_df.index)
    )

    """
    writer = pd.ExcelWriter('output.xlsx')
    df1.to_excel(writer)
    writer.save()
    input("__________________")
    """
    return df1
