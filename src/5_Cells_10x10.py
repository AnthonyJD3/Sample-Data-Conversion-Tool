import pandas as pd
import time

Old_Path = "C:\\Folder\\Old_Sample_File.xlsx"
New_Path = "C:\\Folder\\New_Data_File.xlsx"

df1 = pd.read_excel(Old_Path, header=None)

oldCol0 = df1[0].to_list()
oldCol1 = df1[1].to_list()
oldCol2 = df1[2].to_list()
oldCol3 = df1[3].to_list()
oldCol4 = df1[4].to_list()
oldCol5 = df1[5].to_list()
oldCol6 = df1[6].to_list()
oldCol7 = df1[7].to_list()
oldCol8 = df1[8].to_list()
oldCol9 = df1[9].to_list()

oldCol0 = [oldCol0[i:i+51] for i in range(0, len(oldCol0), 51)]

oldCol0 = [oldCol0[box][i:i+5]+[oldCol0[box][0]] for box in range(len(oldCol0) - 1)
           for i in range(1, len(oldCol0[box]), 5)]

oldCol1 = [oldCol1[i:i+51] for i in range(0, len(oldCol1), 51)]

oldCol1 = [oldCol1[box][i:i+5]+[oldCol1[box][0]] for box in range(len(oldCol1) - 1)
           for i in range(1, len(oldCol1[box]), 5)]

oldCol2 = [oldCol2[i:i+51] for i in range(0, len(oldCol2), 51)]

oldCol2 = [oldCol2[box][i:i+5]+[oldCol2[box][0]] for box in range(len(oldCol2) - 1)
           for i in range(1, len(oldCol2[box]), 5)]

oldCol3 = [oldCol3[i:i+51] for i in range(0, len(oldCol3), 51)]

oldCol3 = [oldCol3[box][i:i+5]+[oldCol3[box][0]] for box in range(len(oldCol3) - 1)
           for i in range(1, len(oldCol3[box]), 5)]

oldCol4 = [oldCol4[i:i+51] for i in range(0, len(oldCol4), 51)]

oldCol4 = [oldCol4[box][i:i+5]+[oldCol4[box][0]] for box in range(len(oldCol4) - 1)
           for i in range(1, len(oldCol4[box]), 5)]

oldCol5 = [oldCol5[i:i+51] for i in range(0, len(oldCol5), 51)]

oldCol5 = [oldCol5[box][i:i+5]+[oldCol5[box][0]] for box in range(len(oldCol5) - 1)
           for i in range(1, len(oldCol5[box]), 5)]

oldCol6 = [oldCol6[i:i+51] for i in range(0, len(oldCol6), 51)]

oldCol6 = [oldCol6[box][i:i+5]+[oldCol6[box][0]] for box in range(len(oldCol6) - 1)
           for i in range(1, len(oldCol6[box]), 5)]

oldCol7 = [oldCol7[i:i+51] for i in range(0, len(oldCol7), 51)]

oldCol7 = [oldCol7[box][i:i+5]+[oldCol7[box][0]] for box in range(len(oldCol7) - 1)
           for i in range(1, len(oldCol7[box]), 5)]

oldCol8 = [oldCol8[i:i+51] for i in range(0, len(oldCol8), 51)]

oldCol8 = [oldCol8[box][i:i+5]+[oldCol8[box][0]] for box in range(len(oldCol8) - 1)
           for i in range(1, len(oldCol8[box]), 5)]

oldCol9 = [oldCol9[i:i+51] for i in range(0, len(oldCol9), 51)]

oldCol9 = [oldCol9[box][i:i+5]+[oldCol9[box][0]] for box in range(len(oldCol9) - 1)
           for i in range(1, len(oldCol9[box]), 5)]

df3 = pd.DataFrame(oldCol0 + oldCol1 + oldCol2 + oldCol3 + oldCol4 + oldCol5 + oldCol6
                   + oldCol7 + oldCol8 + oldCol9, columns="A B C D E Box".split())

df3.loc[:, "Position"] = "Position"

counter = 1

for i in range(len(df3)):
    if i >= 0 and i <= 9:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 10:
        counter = 2

    elif i >= 130 and i <= 139:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 140:
        counter = 3

    elif i >= 260 and i <= 269:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 270:
        counter = 4

    elif i >= 390 and i <= 399:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 400:
        counter = 5

    elif i >= 520 and i <= 529:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 530:
        counter = 6

    elif i >= 650 and i <= 659:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 660:
        counter = 7

    elif i >= 780 and i <= 789:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 790:
        counter = 8

    elif i >= 910 and i <= 919:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 920:
        counter = 9

    elif i >= 1040 and i <= 1049:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

    elif i == 1050:
        counter = 10

    elif i >= 1170 and i <= 1179:
        for j in range(13):
            df3.at[i + (10 * j), 'Position'] = f"{counter}"
        counter += 10

with pd.ExcelWriter(New_Path, engine='openpyxl', mode='a') as writer:
    df3.to_excel(writer, sheet_name='NewDataPositions ' + '{}'.format(time.strftime('%m%d%y-%H%M%S')), index=None)

print('\nA new entry has been added to the Test_New_Data_File file.')
