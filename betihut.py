import os, sys
from docxtpl import DocxTemplate
import pandas as pd




PASSED = 100
TIMESTAMP_STRUCTURE = '%Y-%m-%d %H:%M:%S.%f'
WEAPON_LIST = ['tavor','m4','m16','pistol']
WEAPON_DICT = {'tavor': 'מיקרו תבור','m4': "M4", 'm16': "M16", 'pistol': "אקדח"}


"""Used to translate data from hebrew to english"""
eng_hebrew_dic = {"time_stamp" : "חותמת זמן" ,
"grade": "ניקוד",
"rank": "דרגה",
"platoon":  "מסגרת / פלוגה",
"weapon":  "הנשק עליו אני חותם",
"id":  "מספר אישי",
"first_name":  "שם פרטי",
"last_name":  "שם משפחה",
"initiation_time":  "לפני כמה זמן עברת את ההכשרה על סוג הנשק שאתה מקבל?",
"confidence": "האם לדעתך אתה מספיק בקיא בטיפול והחזקת הנשק שאתה מקבל?",
"last_shooting_range": "לפני כמה זמן ביצעת מטווח בנשק שאתה מקבל?"}

heb_eng_dic = {}
for key, value in eng_hebrew_dic.items():
    heb_eng_dic[value] = key

"""Create pandas dataframe from test_results.xlsx (excel)"""
def load_data(path):
    df = pd.read_excel(path)
    df.rename(columns=heb_eng_dic , inplace=True)
    return df


def rename_weapon(weapon):
    for item in WEAPON_LIST:
        if weapon == WEAPON_DICT[item]: return item
    return weapon


def clean_data(df):
    df = df[heb_eng_dic.values()]
    df = df.loc[df.grade == PASSED]
    df['date'] = df['time_stamp'].apply(lambda x: f"{x.day}/{x.month}/{x.year}")
    df.drop(columns=['time_stamp'], inplace=True)
    df['id'] = df['id'].apply(lambda x: (str(x)).split(".")[0])
    for key in ["initiation_time", 'confidence', 'grade', "last_shooting_range"]:
        df[key] = df[key].apply(lambda x: round(x))
    df["weapon"] = df["weapon"].apply(lambda x: rename_weapon(x))
    return df


"""Create a form from a dictionary of one soldier"""
def create_form(person_dict):
    os.chdir(sys.path[0])
    doc = DocxTemplate('betihut_template.docx') # Reads  the document template from 'betihut_template.docx'
    doc.render(person_dict)
    doc.save(f"{person_dict['id']}.docx") # saves the filled document as {id}.docx


"""Updating a soldier form's dictionary in order to create the form"""
def initiate_dict(dict):
    dict['name'] = dict["first_name"] + " " + dict["last_name"]
    dict[dict["weapon"]] = "V"
    dict[f"A3{dict['confidence']}"] = "V"
    dict[f"A21{dict['initiation_time']}"] = "V"
    dict[f"A22{dict['last_shooting_range']}"] = "V"
    return dict


""""Create a form for each soldier"""
def forms(df):
    for i in range(df.shape[0]):
        personal_dict = initiate_dict(df.iloc[i].to_dict())
        create_form(personal_dict)

