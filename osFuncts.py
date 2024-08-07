import os


def getSubjects():
    years = [folder for folder in os.listdir() if os.path.isdir(folder)]
    years.remove(".idea")
    years.remove("__pycache__")
    subjects = []
    subjectsPath = dict()
    path = os.getcwd()
    for year in years:
        directory = os.getcwd() + "\\" + year
        semesters = os.listdir(directory)
        for semester in semesters:
            for subject in os.listdir(directory + "\\" + semester):
                subjects.append(subject)
                subjectsPath[subject] = year + "\\" + semester + "\\" + subject

    return subjects, subjectsPath

def getCabecalho():
    years = {folder : dict() for folder in os.listdir() if os.path.isdir(folder) and folder!= ".idea" and folder!= "__pycache__"}
    for year in years:
        directory = os.getcwd() + "\\" + year
        years[year] = os.listdir(directory)
    return years

def file_exist_here(file):
    return file in os.listdir()


def file_exist(file, path):
    return file in os.listdir(path)