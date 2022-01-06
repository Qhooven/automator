import PySimpleGUI as sg
import pandas as pd
import pyautogui as pg
import os
import docx
import time
import pathlib
import zipfile
from merge import mergeDiff, getIds, getDocx
from selenium import webdriver
from readexcel import readExcel
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from zipfile import ZipFile


def initGui():
    # Sets up and reads results of initial gui that appears from running main
    firstLayout = [[sg.Button('Update Past Projects')],
                   [sg.Button('Faculty List')]]
    windowMain = sg.Window('Select', firstLayout)
    eventMain, valueMain = windowMain.read()
    windowMain.close()
    if (eventMain == 'Update Past Projects'):
        updateProj()
    else:
        masterFile = sg.popup_get_file('Select all faculty')
        curfFile = sg.popup_get_file('Select curf faculty')
        mergeDiff(curfFile, masterFile)


def updateProj():
    # Update gui. Reads excel file that contains updates to do, opens chrome instance of selenium, prompts login and reads login and authenticate to login to pennintouch.
    updateLayout = [[sg.Text("Username")],
                    [sg.Input()],
                    [sg.Text("Password")],
                    [sg.InputText('', password_char='*')],
                    [sg.Button('Ok')]]
    updateFile = sg.popup_get_file('Select past experiences')
    updatesToDo = readExcel(updateFile)
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": str(
        pathlib.Path(__file__).parent.resolve()) + '\\down'}
    options.add_experimental_option("prefs", prefs)
    path = str(pathlib.Path(__file__).parent.resolve()) + '\\chromedriver.exe'
    driver = webdriver.Chrome(executable_path=path, chrome_options=options)
    driver.get("https://www.curf.upenn.edu/saml_login?destination=user")
    login = '//*[@id="pennname"]'
    passField = '//*[@id="password"]'
    lgnBtn = '//*[@id="login-form"]/div[2]/div/button'
    authField = '//*[@id="penntoken"]'
    authBtn = '//*[@id="formSubmit"]'
    # Create the window
    windowUpdate = sg.Window('Login', updateLayout)
    # Display and interact with the Window
    event, values = windowUpdate.read()
    # Do something with the information gathered
    username = values[0]
    password = values[1]
    driver.find_element_by_xpath(login).send_keys(username)
    driver.find_element_by_xpath(passField).send_keys(password)
    driver.find_element_by_xpath(lgnBtn).click()
    # Finish up by removing from the screen
    windowUpdate.close()
    if (event == 'Ok'):
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
    # Ids gets ids of faculty mentors from faculty list 6 4 2021(currently) as form doesn't require faculty pennids. Getdocx downloads
    # the documents and pictures of the project to the down folder in this file. These are deleted after the updates are complete.
    ids = getIds(updateFile)
    getDocx(updateFile, username, password)
    # Goes through pandas dataframe created from excel file provide. First checks that the profiles of the people and the project itself do not exist.
    # If they exist, nothing happens, otherwise the profiles/project are created.
    for person in updatesToDo:
        truth = checkProfile(driver, person, True)
        print(truth, person[1], person[2])
        truthFacl = checkProfile(driver, person, False)
        print(truthFacl, person[16], person[17])
        if not truth:
            createProfile(driver, person)
        if not truthFacl:
            createProfileFacl(driver, person, ids)
    for person in updatesToDo:
        truthProj = checkProj(driver, person)
        if not truthProj:
            createProj(driver, person)
    # Deletes documents and pictures from directory
    for f in os.listdir(str(pathlib.Path(__file__).parent.resolve()) + '\\down'):
        os.remove(os.path.join(
            str(pathlib.Path(__file__).parent.resolve()) + '\\down', f))


def createProj(driver, person):
    # Creates the project, most inital code just stores xpaths.
    driver.get('https://www.curf.upenn.edu/node/add/project')
    aT = '//*[@id="edit-title"]'
    dT = '//*[@id="edit-field-display-title-und-0-value"]'
    fundingDict = {'Benjamin Franklin Scholars Summer Opportunity Fellowships': 'Benjamin Franklin Scholars Summer Opportunity Fellowships [nid:1673]', 'College Alumni Society Undergraduate Research Grant': 'College Alumni Society Undergraduate Research Grant [nid:1596]', 'Grants for Faculty Mentoring Undergraduate Research': 'Grants for Faculty Mentoring Undergraduate Research [nid:1613]', 'Gelfman International Summer Fund': 'Gelfman International Summer Fund [nid:1682]',
                   'Gutmann-Doyle': 'Gutmann-Doyle Research Opportunities Fund [nid:1614]', 'Hassenfeld Foundation Social Impact Grants': 'Hassenfeld Foundation Social Impact Grants 3 [nid:6821]', 'CLASS OF 1971 ROBERT J. HOLTZ FUND': 'CLASS OF 1971 ROBERT J. HOLTZ FUND [nid:1589]', 'Jumpstart for Juniors': 'Jumpstart for Juniors [nid:1648]', 'Penn Undergraduate Sustainability Action Grant': 'Penn Undergraduate Sustainability Action Grant [nid:1655]',
                   'Penn Undergraduate Research Mentoring Program': 'Penn Undergraduate Research Mentoring Program [nid:1656]', 'Team Grants for Interdisciplinary Activities': 'Team Grants for Interdisciplinary Activities (TGIA) [nid:1658]', 'Vagelos Undergraduate Research Grant': 'Vagelos Undergraduate Research Grant [nid:1668]', 'Association of Alumnae Rosemary D. Mazzatenta Scholars Award': 'Association of Alumnae Rosemary D. Mazzatenta Scholars Award [nid:1672]'}
    body = '//*[@id="edit-body-und-0-value"]'
    select = 'body[und][0][format]'
    stuProf = '//*[@id="edit-field-student-profile-reference-und-0-nid"]'
    faclProf = '//*[@id="edit-field-faculty-profile-reference-und-0-nid"]'
    imgBrowse = '//*[@id="edit-field-image-und-0-browse-button"]'
    imgChoose = '/html/body/div[2]/div/div/div[2]/form/div/div[1]/div/div[1]/input[1]'
    upload = '//*[@id="edit-upload-upload-button"]'
    next = '//*[@id="edit-next"]'
    altText = '//*[@id="edit-field-file-image-alt-text-und-0-value"]'
    saveFrame = '//*[@id="edit-submit"]'
    offering = '//*[@id="edit-field-offering-opportunity-refer-und-0-nid"]'
    submit = '//*[@id="edit-submit"]'
    # Sends basic info, waits a second for people's id to appear/
    stuInp = person[2] + ',' + ' ' + person[1]
    faclInp = person[17] + ',' + ' ' + person[16]
    aTinp = person[2] + ',' + person[1] + ',' + ' ' + person[10]
    offeringInp = fundingDict.get(person[8])
    driver.find_element_by_xpath(aT).send_keys(aTinp)
    driver.find_element_by_xpath(dT).send_keys(person[10])
    driver.find_element_by_xpath(stuProf).send_keys(stuInp)
    time.sleep(1)
    pg.hotkey('down')
    pg.hotkey('enter')
    driver.find_element_by_xpath(faclProf).send_keys(faclInp)
    time.sleep(1)
    pg.hotkey('down')
    pg.hotkey('enter')
    termReceivedDict = {'2014 (Spring)': '//*[@id="edit-field-term-received-und"]/option[8]', 'Fall 2021': '//*[@id="edit-field-term-received-und"]/option[5]',
                        'Spring 2021': '//*[@id="edit-field-term-received-und"]/option[4]', '2014 (Fall)': '//*[@id="edit-field-term-received-und"]/option[2]', '2015 (Spring)': '//*[@id="edit-field-term-received-und"]/option[16]', '2015 (Fall)': '//*[@id="edit-field-term-received-und"]/option[17]', '2016 (Spring)': '//*[@id="edit-field-term-received-und"]/option[8]',
                        '2016 (Fall)': '//*[@id="edit-field-term-received-und"]/option[15]', '2017 (Spring)': '//*[@id="edit-field-term-received-und"]/option[12]', '2017 (Fall)': '//*[@id="edit-field-term-received-und"]/option[13]', '2018 (Spring)': '//*[@id="edit-field-term-received-und"]/option[10]', '2018 (Fall)': '//*[@id="edit-field-term-received-und"]/option[11]', '2019 (Spring)': '//*[@id="edit-field-term-received-und"]/option[7]',
                        '2019 (Fall)': '//*[@id="edit-field-term-received-und"]/option[9]', '2020 (Spring)': '//*[@id="edit-field-term-received-und"]/option[6]', '2020 (Fall)': '//*[@id="edit-field-term-received-und"]/option[3]'}
    driver.find_element_by_xpath(termReceivedDict.get(person[9])).click()
    driver.find_element_by_xpath(offering).send_keys(offeringInp)
    select = Select(driver.find_element_by_name(select))
    select.select_by_index(1)
    # This block adds the project description, from either a docx file or a txt file. Otherwise is not read so most files are docx or txt.
    if isinstance(person[12], str) and person[12] != '':
        if (os.path.splitext(person[12])[1]) == '.docx':
            docxLoc = ''
            for root, dirs, files in os.walk(str(pathlib.Path(__file__).parent.resolve()) + '\\down'):
                if person[12].rsplit('/', 1)[1] in files:
                    docxLoc = root + '\\' + person[12].rsplit('/', 1)[1]
            doc = docx.Document(docxLoc)
            fullText = []
            for para in doc.paragraphs:
                fullText.append(para.text)
            for x in fullText:
                driver.find_element_by_xpath(body).send_keys(x)
                driver.find_element_by_xpath(body).send_keys(Keys.RETURN)
        elif os.path.splitext(person[12])[1] == '.txt':
            txtLoc = ''
            for root, dirs, files in os.walk(str(pathlib.Path(__file__).parent.resolve()) + '\\down'):
                if person[12].rsplit('/', 1)[1] in files:
                    txtLoc = root + '\\' + person[12].rsplit('/', 1)[1]
            f = open(txtLoc, 'r')
            txt = f.read()
            driver.find_element_by_xpath(body).send_keys(txt)
            driver.find_element_by_xpath(body).send_keys(Keys.RETURN)
        else:
            pass
    # Adds vimeo/youtube link
    if isinstance(person[11], str):
        driver.find_element_by_xpath(body).send_keys(person[11])
    # Gets photo location in computer if not zip, if zip does that but selects first file in the zip folder.
    photoloc = ''
    if (os.path.splitext(person[14])[1]) != '.zip':
        for root, dirs, files in os.walk(str(pathlib.Path(__file__).parent.resolve()) + '\\down'):
            if person[14].rsplit('/', 1)[1] in files:
                photoloc = root + '\\' + person[14].rsplit('/', 1)[1]
    else:
        for root, dirs, files in os.walk(str(pathlib.Path(__file__).parent.resolve()) + '\\down'):
            if person[14].rsplit('/', 1)[1] in files:
                zip = ZipFile(str(pathlib.Path(__file__).parent.resolve(
                )) + '\\down' + '\\' + person[14].rsplit('/', 1)[1])
                for x in zip.namelist():
                    if 'MACOS' not in x:
                        photoloc = root + '\\' + x
                        with zipfile.ZipFile(str(pathlib.Path(__file__).parent.resolve()) + '\\down' + '\\' + person[14].rsplit('/', 1)[1], 'r') as zip_ref:
                            zip_ref.extractall(
                                str(pathlib.Path(__file__).parent.resolve()) + '\\down')
                        zip.close()
                        zip_ref.close()
                        break
    # Adds picture, sets up gui so user can manually crop photo.
    driver.find_element_by_xpath(imgBrowse).click()
    driver.switch_to.frame(
        driver.find_element_by_xpath('//*[@id="mediaBrowser"]'))
    time.sleep(2)
    driver.find_element_by_xpath(imgChoose).send_keys(photoloc)
    driver.find_element_by_xpath(upload).click()
    time.sleep(1)
    driver.find_element_by_xpath(next).click()
    while(True):
        cropLayout = [[sg.Text("Manually Crop Photo, Select ok when done")],
                      [sg.Button('Ok')]]
        cropWindow = sg.Window('Login', cropLayout)
        event, values = cropWindow.read()
        if event == 'Ok':
            cropWindow.close()
            break

        else:
            pass
    driver.find_element_by_xpath(altText).send_keys(
        person[1] + ' ' + person[2])
    driver.find_element_by_xpath(saveFrame).click()
    driver.switch_to.default_content()
    # Comment out unless testing pls
    pubOptions = '//*[@id="project-node-form"]/div/div[13]/ul/li[6]'
    pubBox = '//*[@id="edit-status"]'
    driver.find_element_by_xpath(pubOptions).click()
    driver.find_element_by_xpath(pubBox).click()
    driver.find_element_by_xpath(submit).click()


def createProfile(driver, person):
    # Mostly xpaths
    driver.get('https://www.curf.upenn.edu/node/add/profile')
    aT = '//*[@id="edit-title"]'
    firstName = '//*[@id="edit-field-first-name-und-0-value"]'
    lastName = '//*[@id="edit-field-last-name-und-0-value"]'
    stuPath = '//*[@id="field_profile_role"]/option[10]'
    add = '//*[@id="field_profile_role"]/a'
    stuYear = '//*[@id="edit-field-student-year-und-0-value"]'
    majorDict = {'Africana Studies': '//*[@id="field_academic_major"]/option[1]', 'Ancient History': '//*[@id="field_academic_major"]/option[2]', 'Anthropology': '//*[@id="field_academic_major"]/option[3]', 'Architecture': '//*[@id="field_academic_major"]/option[4]', 'Asian American Studies': '//*[@id="field_academic_major"]/option[5]',
                 'Biochemistry': '//*[@id="field_academic_major"]/option[6]', 'Bioengineering': '//*[@id="field_academic_major"]/option[7]', 'Biological Mathematics': '//*[@id="field_academic_major"]/option[8]', 'Biology': '//*[@id="field_academic_major"]/option[9]',
                 'Biomedical Science': '//*[@id="field_academic_major"]/option[10]', 'Biophysics': '//*[@id="field_academic_major"]/option[11]', 'Business Analytics': '//*[@id="field_academic_major"]/option[12]', 'Business Economics and Public Policy': '//*[@id="field_academic_major"]/option[13]', 'Chemical and Biomolecular Engineering': '//*[@id="field_academic_major"]/option[14]',
                 'Chemistry': '//*[@id="field_academic_major"]/option[15]', 'Cinema Studies': '//*[@id="field_academic_major"]/option[16]', 'Classical Studies': '//*[@id="field_academic_major"]/option[17]', 'Cognitive Science': '//*[@id="field_academic_major"]/option[18]', 'Communication': '//*[@id="field_academic_major"]/option[19]',
                 'Comparative Literature': '//*[@id="field_academic_major"]/option[20]', 'Computational Biology': '//*[@id="field_academic_major"]/option[21]', 'Computer Engineering': '//*[@id="field_academic_major"]/option[22]', 'Computer Science': '//*[@id="field_academic_major"]/option[23]', 'Criminology': '//*[@id="field_academic_major"]/option[24]', 'Digital Media Design': '//*[@id="field_academic_major"]/option[25]', 'Earth Science': '//*[@id="field_academic_major"]/option[26]',
                 'East Asian Area Studies': '//*[@id="field_academic_major"]/option[27]', 'East Asian Languages and Civilizations': '//*[@id="field_academic_major"]/option[28]', 'Economics': '//*[@id="field_academic_major"]/option[29]',
                 'Electrical Engineering': '//*[@id="field_academic_major"]/option[30]', 'English': '//*[@id="field_academic_major"]/option[31]', 'Environmental Studies': '//*[@id="field_academic_major"]/option[32]', 'Finance': '//*[@id="field_academic_major"]/option[33]', 'Fine Arts': '//*[@id="field_academic_major"]/option[34]', 'French and Francophone Studies': '//*[@id="field_academic_major"]/option[35]', "Gender, Sexuality, and Women's Studies": '//*[@id="field_academic_major"]/option[36]',
                 'German': '//*[@id="field_academic_major"]/option[37]', 'Health and Societies': '//*[@id="field_academic_major"]/option[38]', 'Hispanic Studies': '//*[@id="field_academic_major"]/option[39]', 'History': '//*[@id="field_academic_major"]/option[40]', 'History of Art': '//*[@id="field_academic_major"]/option[41]', 'International Relations': '//*[@id="field_academic_major"]/option[42]', 'International Studies': '//*[@id="field_academic_major"]/option[43]',
                 'Italian Studies': '//*[@id="field_academic_major"]/option[44]', 'Jewish Studies': '//*[@id="field_academic_major"]/option[45]', 'Latin American and Latino Studies': '//*[@id="field_academic_major"]/option[46]', 'Linguistics': '//*[@id="field_academic_major"]/option[47]', 'Logic, Information and Computation': '//*[@id="field_academic_major"]/option[48]', 'Management': '//*[@id="field_academic_major"]/option[49]', 'Marketing': '//*[@id="field_academic_major"]/option[50]',
                 'Materials Science and Engineering': '//*[@id="field_academic_major"]/option[51]', 'Mathematical Economics': '//*[@id="field_academic_major"]/option[52]', 'Mathematics': '//*[@id="field_academic_major"]/option[53]', 'Mechanical Engineering and Applied Mechanics': '//*[@id="field_academic_major"]/option[54]', 'Medical Anthropology and Global Health': '//*[@id="field_academic_major"]/option[55]', 'Modern Middle Eastern Studies': '//*[@id="field_academic_major"]/option[56]',
                 'Molecular and Cell Biology': '//*[@id="field_academic_major"]/option[57]', 'Music': '//*[@id="field_academic_major"]/option[58]', 'Near Eastern Languages and Civilizations': '//*[@id="field_academic_major"]/option[59]', 'Networked and Social Systems Engineering': '//*[@id="field_academic_major"]/option[60]', 'Neuroscience': '//*[@id="field_academic_major"]/option[61]', 'Nursing': '//*[@id="field_academic_major"]/option[62]', 'Organizational Dynamics': '//*[@id="field_academic_major"]/option[63]',
                 'Philosophy': '//*[@id="field_academic_major"]/option[64]', 'Philosophy, Politics and Economics': '//*[@id="field_academic_major"]/option[65]', 'Physics': '//*[@id="field_academic_major"]/option[66]', 'Political Science': '//*[@id="field_academic_major"]/option[67]', 'Psychology': '//*[@id="field_academic_major"]/option[68]', 'Religious Studies': '//*[@id="field_academic_major"]/option[69]', 'Romance Languages': '//*[@id="field_academic_major"]/option[70]', 'Russian': '//*[@id="field_academic_major"]/option[71]',
                 'Science, Technology and Society': '//*[@id="field_academic_major"]/option[72]', 'Sociology': '//*[@id="field_academic_major"]/option[73]', 'South Asia Studies': '//*[@id="field_academic_major"]/option[74]', 'Spanish': '//*[@id="field_academic_major"]/option[75]', 'Systems Science and Engineering': '//*[@id="field_academic_major"]/option[76]', 'Theatre Arts': '//*[@id="field_academic_major"]/option[77]', 'Undecided': '//*[@id="field_academic_major"]/option[78]', 'Urban Studies': '//*[@id="field_academic_major"]/option[79]',
                 'Visual Studies': '//*[@id="field_academic_major"]/option[80]'}
    majorAdd = '//*[@id="field_academic_major"]/a'
    schoolDict = {'Annenberg': '//*[@id="field_school_reference"]/option[1]', 'Arts and Sciences/Graduate': '//*[@id="field_school_reference"]/option[2]', 'College': '//*[@id="field_school_reference"]/option[3]', 'Dental Medicine': '//*[@id="field_school_reference"]/option[4]', 'Design': '//*[@id="field_school_reference"]/option[5]', 'Education': '//*[@id="field_school_reference"]/option[6]', 'Engineering & Applied Sciences': '//*[@id="field_school_reference"]/option[7]', 'Law': '//*[@id="field_school_reference"]/option[8]',
                  'Liberal & Professional Studies': '//*[@id="field_school_reference"]/option[9]', 'Medicine': '//*[@id="field_school_reference"]/option[10]', 'Nursing': '//*[@id="field_school_reference"]/option[11]', 'Social Policy and Practice': '//*[@id="field_school_reference"]/option[12]', 'Veterinary Medicine': '//*[@id="field_school_reference"]/option[13]', 'Wharton': '//*[@id="field_school_reference"]/option[14]'}
    schoolAdd = '//*[@id="field_school_reference"]/a'
    pennId = '//*[@id="edit-field-pennid-und-0-value"]'
    major = majorDict.get(person[5])
    school = schoolDict.get(person[4])
    save = '//*[@id="edit-submit"]'
    pennIdStr = str(person[0])
    # Sends basic info like Name, major
    aTInput = person[2] + ',' + ' ' + person[1]
    driver.find_element_by_xpath(aT).send_keys(aTInput)
    driver.find_element_by_xpath(firstName).send_keys(person[1])
    driver.find_element_by_xpath(lastName).send_keys(person[2])
    driver.find_element_by_xpath(stuPath).click()
    driver.find_element_by_xpath(add).click()
    if(majorDict.get(person[5])):
        driver.find_element_by_xpath(major).click()
        driver.find_element_by_xpath(majorAdd).click()
    driver.find_element_by_xpath(school).click()
    driver.find_element_by_xpath(schoolAdd).click()
    driver.find_element_by_xpath(pennId).send_keys(pennIdStr)
    # Comment out unless testing pls
    driver.find_element_by_xpath(save).click()


def createProfileFacl(driver, person, ids):
    # Same as above, however gets ids from the get ids functions
    driver.get('https://www.curf.upenn.edu/node/add/profile')
    aT = '//*[@id="edit-title"]'
    firstName = '//*[@id="edit-field-first-name-und-0-value"]'
    lastName = '//*[@id="edit-field-last-name-und-0-value"]'
    faclPath = '//*[@id="field_profile_role"]/option[7]'
    add = '//*[@id="field_profile_role"]/a'
    faclEmail = '//*[@id="edit-field-contact-email-address-und-0-email"]'
    idInput = '//*[@id="edit-field-pennid-und-0-value"]'
    save = '//*[@id="edit-submit"]'
    aTInput = person[17] + ',' + ' ' + person[16]
    id = 0
    for people in ids:
        if person[16].lower() + ' ' + person[17].lower() == people[0].lower() + ' ' + people[1].lower():
            id = int(people[2])
            break
    driver.find_element_by_xpath(aT).send_keys(aTInput)
    driver.find_element_by_xpath(firstName).send_keys(person[16])
    driver.find_element_by_xpath(lastName).send_keys(person[17])
    driver.find_element_by_xpath(faclPath).click()
    driver.find_element_by_xpath(add).click()
    driver.find_element_by_xpath(faclEmail).send_keys(person[20])
    if (id != 0):
        idStr = str(id)
        driver.find_element_by_xpath(idInput).send_keys(idStr)
    # Comment out unless testing pls
    driver.find_element_by_xpath(save).click()


def checkProfile(driver, person, student):
    # Checks if student/faculty profile exists by search and if the search yields something comparing it to the person's name
    driver.get('https://www.curf.upenn.edu/search/site')
    profile = '//*[@id="block-system-main"]/div/div/ol/li[1]/h3/a'
    searchBar = '//*[@id="edit-keys-1"]'
    searchBtn = '//*[@id="edit-submit-1"]'
    if(student):
        name = person[1] + ' ' + person[2]
    else:
        name = person[16] + ' ' + person[17]
    driver.find_element_by_xpath(searchBar).send_keys(name)
    driver.find_element_by_xpath(searchBtn).click()
    try:
        actName = driver.find_element_by_xpath(profile).text
    except(NoSuchElementException):
        return False
    if (name == actName):
        return True
    else:
        return False


def checkProj(driver, person):
    # Checks if project exists in similar fashion to the way a person's existence is checked.
    driver.get('https://www.curf.upenn.edu/research/need-to-know/past-projects')
    searchBar = '//*[@id="edit-text"]'
    searchBtn = '//*[@id="edit-submit-past-projects"]'
    project = '//*[@id="block-system-main"]/div/div/div/div[2]/div/div/div/div/div[2]/table/tbody/tr/td[1]/a'
    driver.find_element_by_xpath(searchBar).send_keys(person[10])
    driver.find_element_by_xpath(searchBtn).click()
    time.sleep(8)
    try:
        actName = driver.find_element_by_xpath(project).text
    except(NoSuchElementException):
        return False
    if (person[10] == actName):
        return True
    else:
        return False
