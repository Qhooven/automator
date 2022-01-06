import pandas as pd


def readExcel(excel):
    df = pd.read_excel(excel)
    people = []
    for index, row in df.iterrows():
        person = []
        for x in range(0, 21):
            person.append(df.iloc[index][x])
        people.append(person)
    return people
