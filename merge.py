from PySimpleGUI.PySimpleGUI import Listbox
import pandas as pd
import PySimpleGUI as sg
import os
import pathlib
import time
from zipfile import ZipFile
import zipfile
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def mergeDiff(loc1, loc2):
    rdEntries = pd.read_excel(loc1)
    facultyList = pd.read_excel(loc2)
    rdEntries.dropna()
    #rdEntries['Penn Id (EG)'] = rdEntries['Penn Id (EG)'].astype('int64')
    merged = rdEntries.merge(facultyList, how='outer',
                             indicator='source', on='Penn Id (EG)')
    merged = merged[merged['source'] == 'right_only']
    merged.drop(merged.filter(regex='_x$').columns.tolist(),
                axis=1, inplace=True)
    merged.to_excel('OnlyOnSecond.xlsx')


def getIds(loc1):
    rdEntries = pd.read_excel(loc1)
    facultyList = pd.read_excel(
        str(pathlib.Path(__file__).parent.resolve()) + '\\Faculty List 6.4.21.xlsx')
    rdEntries = rdEntries.drop(rdEntries.columns[[
                               0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 18, 19, 20]], axis=1)
    facultyList = facultyList.drop(
        facultyList.columns[[1, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]], axis=1)
    rdEntries["First Name (EG)"] = rdEntries["First Name (EG)"].str.lower()
    rdEntries["Last Name (EG)"] = rdEntries["Last Name (EG)"].str.lower()
    facultyList["First Name (EG)"] = facultyList["First Name (EG)"].str.lower()
    facultyList["Last Name (EG)"] = facultyList["Last Name (EG)"].str.lower()
    merged = rdEntries.merge(facultyList, how='outer', indicator='source', on=[
                             'First Name (EG)', 'Last Name (EG)'])
    merged = merged[merged['source'] == 'both']
    merged = merged.drop_duplicates()
    vals = merged.to_numpy()
    return vals


def getDocx(pastexperiences, username, password):
    blah = pd.read_excel(pastexperiences)
    people = []
    for index, row in blah.iterrows():
        person = []
        for x in range(0, 21):
            person.append(blah.iloc[index][x])
        people.append(person)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    prefs = {"download.default_directory": str(
        pathlib.Path(__file__).parent.resolve()) + '\\down'}
    options.add_experimental_option("prefs", prefs)
    path = str(pathlib.Path(__file__).parent.resolve()) + '\\chromedriver.exe'
    driver = webdriver.Chrome(executable_path=path, chrome_options=options)
    driver.get(people[0][12])
    lgn = '//*[@id="content"]/div/p[2]/a'
    login = '//*[@id="pennname"]'
    passField = '//*[@id="password"]'
    lgnBtn = '//*[@id="login-form"]/div[2]/div/button'
    authField = '//*[@id="penntoken"]'
    authBtn = '//*[@id="formSubmit"]'
    body = '/html/body/pre'
    driver.find_element_by_xpath(lgn).click()
    driver.find_element_by_xpath(login).send_keys(username)
    driver.find_element_by_xpath(passField).send_keys(password)
    driver.find_element_by_xpath(lgnBtn).click()
    authLayout = [
        [sg.Button('Used push notifcation')]]
    authLayout = [[sg.Text("Enter authentification code")],
                  [sg.Input()],
                  [sg.Button('Ok')],
                  [sg.Button('Used push notifcation')]]
    authWindow = sg.Window('Authentification', authLayout)
    event, values = authWindow.read()
    if (event == 'Ok'):
        authCode = values[0]
        driver.find_element_by_xpath(authField).send_keys(authCode)
        driver.find_element_by_xpath(authBtn).click()
        authWindow.close()
    authWindow.close()
    path = str(pathlib.Path(__file__).parent.resolve()) + '\\down'
    for person in people:
        driver.get(person[14])
        if os.path.splitext(person[14])[1] == '.zip':
            time.sleep(2)
        else:
            driver.save_screenshot(str(pathlib.Path(__file__).parent.resolve(
            )) + '\\down' + '\\' + person[14].rsplit('/', 1)[1])
        if isinstance(person[12], str):
            driver.get(person[12])
            if (os.path.splitext(person[12])[1]) == '.txt':
                text = driver.find_element_by_xpath(body).text
                completeName = person[12].rsplit('/', 1)[1]
                f = open(os.path.join(path, completeName), "w")
                f.write(text)
                f.close()
    driver.close()


def zip(pastexperiences):
    blah = pd.read_excel(pastexperiences)
    people = []
    for index, row in blah.iterrows():
        person = []
        for x in range(0, 21):
            person.append(blah.iloc[index][x])
        people.append(person)
    for root, dirs, files in os.walk(str(pathlib.Path(__file__).parent.resolve()) + '\\down'):
        if person[14].rsplit('/', 1)[1] in files:
            zip = ZipFile(str(pathlib.Path(__file__).parent.resolve()) +
                          '\\down' + '\\' + person[14].rsplit('/', 1)[1])
            # print(zip.namelist()[0])
            for x in zip.namelist():
                if 'MACOS' not in x:
                    print(x)
            print(str(pathlib.Path(__file__).parent.resolve()) +
                  '\\down' + person[14].rsplit('/', 1)[1])
            with zipfile.ZipFile(str(pathlib.Path(__file__).parent.resolve()) + '\\down' + '\\' + person[14].rsplit('/', 1)[1], 'r') as zip_ref:
                zip_ref.extractall(
                    str(pathlib.Path(__file__).parent.resolve()) + '\\down')


def stats(pastexperiences):
    blah = pd.read_excel(pastexperiences)
    people = []
    for index, row in blah.iterrows():
        person = []
        for x in range(0, 21):
            person.append(blah.iloc[index][x])
        people.append(person)
    jpg = 0
    png = 0
    otherPic = list()
    docx = 0
    txt = 0
    otherTxt = list()
    for person in people:
        if (os.path.splitext(person[12])[1]) == '.txt':
            txt = txt + 1
        elif (os.path.splitext(person[12])[1]) == '.docx':
            docx = docx + 1
        else:
            otherTxt.append((os.path.splitext(person[12])[1]))
        if (os.path.splitext(person[14])[1]) == '.jpg' or (os.path.splitext(person[14])[1]) == '.jpeg':
            jpg = jpg + 1
        elif (os.path.splitext(person[14])[1]) == '.png':
            png = png + 1
        else:
            otherPic.append((os.path.splitext(person[14])[1]))
    print('jpg', ':', jpg, 'png', ':', png, 'other', ':', otherPic,
          'docx', ':', docx, 'txt', ':', txt, 'otherTxt', ':', otherTxt)
