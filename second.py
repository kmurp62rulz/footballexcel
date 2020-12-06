from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

wb = load_workbook('test.xlsx')
main = wb['main']
max = len(main['A']) + 1

season = input('What season are you in? \n')


if season in wb.sheetnames:
    print('Match found')
    ws = wb[season]

else:
    if input('No matches found, create new sheet? (y/n) \n') == 'y':
        ws = wb.create_sheet(season)
        
        for name in range(1, max):
            ws.cell(row=name,column=1).value = main.cell(row=name,column=1).value
        names = ws.cell(row=1, column=1)
        names.font = Font(bold=True)
        names.value = 'Names'
        print('New season: ' + str(season) + ' created successfully!')
    else:
        print('ok')

nextcol = ws.max_column + 1
for name in range(2, max):
    #ask user how many goals they scored this week
    goals = input('How many goals did ' + str(ws.cell(row=name,column=1).value) + ' score this week? \n')
    #create new cell at len(ws)+1
    newcell = ws.cell(row=name, column=nextcol)
    newcell.font = Font(bold=False)
    newcell.value = int(goals)
    week = ws.cell(row=1, column=nextcol)
    week.font = Font(bold=True)
    week.value = 'Week ' + str(nextcol - 1)


#for week in range(1, ws.max_column):
#    print(week)
#    #make titles with week + len(ws)
#    nextnamecol = 2
#    titlecell = ws.cell(row=1, column=nextnamecol)
#    titlecell.font = Font(bold=True)
#    titlecell.value = 'Week ' + str(nextnamecol - 1)
#    nextnamecol += 1


wb.save('test.xlsx')
    