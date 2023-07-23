import pandas as pd
import os
import json
import xlsxwriter
import math

# Creates pace doc if it doesn't exist.
def paceDocCreation(path):
    worksheet.set_column(0, 0, 8)
    worksheet.set_column(1, 14, 12)
    worksheet.set_column(15, 15, 30)

    boldBot = workbook.add_format({'bold': True, 'bottom': 5})

    # Writing titles.
    titles = ["Run", "Date", "Nether Enter",
            "Bastion Enter", "Fort Enter", "Blind",
            "Stronghold", "Enter End", "Kill Dragon",
            "Bastion Travel", "Bastion Split", "Fort Split",
            "SH Travel", "SH Split", "End Split", "File Name"]
    for i in range(len(titles)):
        worksheet.write(0, i, titles[i], boldBot)


def writeAverageBest(startRow, endRow):
    boldRight = workbook.add_format({'num_format': 'mm:ss', 'bold': True, 'right': 5})
    bold = workbook.add_format({'num_format': 'mm:ss', 'bold': True})

    # Writing average formulae.
    worksheet.write(startRow, 0, "Averages", bold)
    worksheet.write(startRow, 1, "N/A", boldRight)
    for i in range(2, 15):
        if i in [8, 14]:
            worksheet.write(startRow, i, f'=MEDIAN({chr(i+65).upper()}{startRow + 3}:{chr(i+65).upper()}{endRow})', boldRight)
        else:
            worksheet.write(startRow, i, f'=MEDIAN({chr(i+65).upper()}{startRow + 3}:{chr(i+65).upper()}{endRow})', bold)
    worksheet.write(startRow, 15, "N/A", bold)

    # Writing best formulae.
    worksheet.write(startRow + 1, 0, "PBs", bold)
    worksheet.write(startRow + 1, 1, "N/A", boldRight)
    for i in range(2, 15):
        if i in [8, 14]:
            worksheet.write(startRow + 1, i, f'=MIN({chr(i+65).upper()}{startRow + 3}:{chr(i+65).upper()}{endRow})', boldRight)
        else:
            worksheet.write(startRow + 1, i, f'=MIN({chr(i+65).upper()}{startRow + 3}:{chr(i+65).upper()}{endRow})', bold)
    worksheet.write(startRow + 1, 15, "N/A", bold)

# Checks if pace doc exists.
def paceDocCheck(path):
    if not os.path.exists(path):
        pass
    paceDocCreation(path)

# Checks if a run hasn't already been added to the excel doc already.
def newFileCheck(path):
    return True

# Checks if a run made it into the nether.
def reachedNetherCheck(data):
    if any("enter_nether" in d.values() for d in data["timelines"]):
        return True
    return False

def rsgCheck(data, rsg):
    if rsg and "Speedrun" in data["world_name"]:
            return True
    if not rsg and "mcsrranked" in data["world_name"]:
        return True
    return False

# Reads the run writes its details to the worksheet.
def readJson(f, currentRow, rsg):
    j = open(f)
    data = json.load(j)
    if rsgCheck(data, rsg) and reachedNetherCheck(data):
        print(f)
        currentCol = 0

        timeBgs = ['00EEFF', '00DDEE', '00CCDD', '00BBCC', '00AABB', 'FF9999']

        # Write run number. Is xl index - 3.
        worksheet.write_number(currentRow, currentCol, currentRow - 2)
        currentCol += 1

        # Write date of run in form dd/mm/yy.
        worksheet.write(currentRow, 1, data["date"]/1000/60/60/24 + 25569, format1)
        currentCol += 1

        # Write time of timelines.
        bastionFound = False
        fortFirst = False
        for i in data["timelines"]:
            if i["name"] not in ["found_villager", "portal_no_1", "nether_travel_blind", "nether_travel_home"]:
                if i["name"] == "enter_bastion":
                    bastionFound = True
                elif i["name"] == "enter_fortress" and not bastionFound:
                    fortFirst = True
                if fortFirst and i["name"] in ["enter_fortress", "enter_bastion"]:
                    if i["igt"]/1000/60/60 < 1:
                        worksheet.write(currentRow, currentCol, i["igt"]/1000/60/60/24, format2Fort)
                    else:
                        worksheet.write(currentRow, currentCol, i["igt"]/1000/60/60/24, format3Fort)
                elif i["name"] == "enter_nether":
                        minute = math.floor(i["igt"]/1000/60) - 1
                        if minute > 4:
                            minute = 4
                        if not rsg:
                            minute += 1
                        formatSubX = workbook.add_format({'bg_color': timeBgs[minute], 'num_format': 'mm:ss', 'align': 'right'})
                        worksheet.write(currentRow, currentCol, i["igt"]/1000/60/60/24, formatSubX)
                else:
                    if i["igt"]/1000/60/60 < 1:
                        worksheet.write(currentRow, currentCol, i["igt"]/1000/60/60/24, format2)
                    else:
                        worksheet.write(currentRow, currentCol, i["igt"]/1000/60/60/24, format3)
                currentCol += 1

        # Write reset if event not reached.
        if rsg:
            msg = "Reset"
        else:
            msg = "Forfeit/Loss"
        for i in range(currentCol, 9):
            worksheet.write(currentRow, currentCol, msg, format4)
            currentCol += 1

        # Write splits.
        worksheet.write(currentRow, currentCol, f'=IFERROR({chr(68).upper()}{currentRow + 1}-{chr(67).upper()}{currentRow + 1}, "Reset")', format5)
        currentCol += 1
        for i in range(69, 74):
            worksheet.write(currentRow, currentCol, f'=IFERROR({chr(i).upper()}{currentRow + 1}-{chr(i-1).upper()}{currentRow + 1}, "Reset")', format2)
            currentCol += 1

        # Write filename.
        worksheet.write(currentRow, currentCol, f, format6)
        j.close()
        return True
    j.close()
    return False

def appendToDf(df, row):
    df.loc[len(df)] = row

def appendToPaces(df, path):
    df2 = pd.read_excel(path)
    df3 = pd.concat([df2, df])
    df3.to_excel("mypaces.xlsx", index=False)

path = r".\paces.xlsx"
directory = rf"{os.environ['USERPROFILE']}\speedrunigt\records"

workbook = xlsxwriter.Workbook(path)

# Workbook formats.
format1 = workbook.add_format({'num_format': 'dd/mm/yy', 'right': 5, 'align': 'right'})
format2 = workbook.add_format({'num_format': 'mm:ss', 'align': 'right'})
format3 = workbook.add_format({'num_format': '[h]:mm:ss', 'align': 'right'})
format2Fort = workbook.add_format({'bg_color': '#FFD700', 'num_format': 'mm:ss', 'align': 'right'})
format3Fort = workbook.add_format({'bg_color': '#FFD700', 'num_format': '[h]:mm:ss', 'align': 'right'})
format4 = workbook.add_format({'bg_color': '#AAAAAA', 'align': 'right'})
format5 = workbook.add_format({'num_format': 'mm:ss', 'left': 5, 'align': 'right'})
format6 = workbook.add_format({'left': 5})

worksheet = workbook.add_worksheet()
paceDocCheck(path)
files = os.listdir(directory)
files = [os.path.join(directory, filename) for filename in files]
files.sort(key=os.path.getctime)
curRow = 3
for f in files:
    if os.path.isfile(f) and f.endswith('json'):
        if newFileCheck(f):
            if (readJson(f, curRow, True)):
                curRow += 1

split = curRow
writeAverageBest(1, split)
curRow += 2

for f in files:
    if os.path.isfile(f) and f.endswith('json'):
        if newFileCheck(f):
            if (readJson(f, curRow, False)):
                curRow += 1
writeAverageBest(split, curRow)
workbook.close()