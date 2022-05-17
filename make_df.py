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
