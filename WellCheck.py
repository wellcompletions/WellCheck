from openpyxl import load_workbook
import sys 

if __name__ == "__main__":
    
    with open(sys.argv[1].rstrip(".xlsx")+'_ERRORS.txt', 'w') as f:

        #change these to match the Perf sheet cells    
        cellToe = 'AF22'      
        cellHeel = 'AF26'

        workbookPerf = load_workbook(filename=sys.argv[1], read_only=True, data_only=True)
        # grab number of stages from the sheet tab 
        numstages = int(workbookPerf.sheetnames[1][0:2:])
        
        # print the header for both the terminal and the file 
        print('\n\nLoading Perforations ...\n')
        print('Perforations', file=f)
        print('\n',sys.argv[1][2::],'\n',workbookPerf.sheetnames[1], file=f)
        print(sys.argv[1][2::],'\n')
        print(workbookPerf.sheetnames[1])
        # pick the perf sheet with stages
        sheet = workbookPerf[workbookPerf.sheetnames[1]]
        perfDepth = []  # list because it is mutable
        for i, row in enumerate(sheet.iter_rows(min_row=8,
                                    max_row=(numstages*2)+7,
                                    min_col=3,
                                    max_col=15,
                                    values_only=True)):
            # pick every other row (evens)
            if i % 2 == 0:
                for cell in row:  
                    if isinstance(cell, float) or isinstance(cell, int):          
                        perfDepth.append(round(cell))   
                        print(repr(round(cell)).rjust(5), end = ' ')
                        print(repr(round(cell)).rjust(5), end = ' ', file=f)
                print()
                print(file=f)
        # grab deepest and shallowest perfs
        deepPerf = perfDepth[0]
        shallowPerf = perfDepth[len(perfDepth)-1]
        
        
        print('\nLoading Collars ...')
        print()
        workbookCollars = load_workbook(filename=sys.argv[2], read_only=True, data_only=True)
        
        print(sys.argv[2][2::].strip(), workbookCollars.sheetnames[1],'\n')
        print('\n', sys.argv[2][2::].strip(), workbookCollars.sheetnames[1],'\n', file=f)
        # select the collars sheet
        sheetCollar = workbookCollars[workbookCollars.sheetnames[1]]
        collarDepth = []
        collarCount = 0
        for k, row in enumerate(sheetCollar.iter_rows(min_row=12,
                                    max_row=500,
                                    min_col=21,
                                    max_col=21,
                                    values_only=True), start=1):
            for cell in row:
                if isinstance(cell, float) or isinstance(cell, int):
                    collarCount +=1
                    collarDepth.append(round(cell))   
                    print(repr(round(cell)).rjust(5), end=' ')
                    print(repr(round(cell)).rjust(5), end=' ', file=f)
                    if k % 13 == 0:
                        print()
                        print(file=f)
        print('\n\n', collarCount, 'Collars found.')   
        print('\n\n', collarCount, 'Collars found.', file = f)   
        print('\nList of conflicts: \n')
        print('Collar      Perf \n', end='')
        print('------     ------ \n', end='')
        print('\nList of conflicts: \n', file=f)
        print('Collar      Perf \n', end='',file=f)
        print('------     ------ \n', end='',file=f)

        conflict = []  # create a conflict list for some project later
        for collar in tuple(collarDepth):
            
            for perf in tuple(perfDepth):
                if int(collar) == int(perf):
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is the same')
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is the same',file = f)
                    conflict.append(perf)
                elif int(collar)+1 == int(perf):
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is +1 above')
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is +1 above',file = f)
                    conflict.append(perf)
                elif int(collar)+2 == int(perf):
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is +2 above')
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is +2 above',file = f)
                    conflict.append(perf)
                elif int(collar)-1 == int(perf):
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is -1 below')
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is -1 below',file = f)
                    conflict.append(perf)
                elif int(collar)-2 == int(perf):
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is -2 below')
                    print(repr(collar).rjust(5), '    ', repr(perf).rjust(5), ' is -2 below',file = f) 
                    conflict.append(perf)
        if len(conflict) == 0:
            print()
            print(file = f)
            print('No conflicts detected.')
            print('No conflicts detected.', file = f)
            print()
            print(file = f)
        else:
            print()
            print(file = f)
            print(len(conflict),'conflicts detected.')
            print(len(conflict),'conflicts detected.', file = f)
            print()
            print(file = f)

        print('Loading Survey ...')
        print()
        print('Survey', file = f)
        print(file = f)
        print(sys.argv[3][2::].strip())
        print('\n',sys.argv[2][2::].strip(), file = f)
        workbookSurvey = load_workbook(filename=sys.argv[3], read_only=True, data_only=True)
        sheetSurvey = workbookSurvey[workbookSurvey.sheetnames[0]]
        surveyKB = 0
        surveyHeel = 0 
        surveyToe = 0
        print(sheetSurvey.cell(8,9).value[3:5], 'ft KB depth from header found')
        for i, row in enumerate(sheetSurvey.iter_rows(min_row=24,
                                    max_row=300,
                                    min_col=0,
                                    max_col=2,
                                    values_only=True)):
            for cell in row:
                if repr(cell)[1:3] == "KB":
                    surveyKB = int(round(row[1]))
                    print(round(row[1]), 'ft KB depth from data found')
                    print(round(row[1]), 'ft KB depth from data found', file = f)
                elif repr(cell)[1:19] == "Cross Setback Heel":
                    surveyHeel = round(row[1])   
                    print(surveyHeel, 'Cross Setback Heel depth found') 
                    print(surveyHeel, 'Cross Setback Heel depth found', file = f)
                elif repr(cell)[1:18] == "Cross Setback Toe":
                    surveyToe = round(row[1])
                    print(surveyToe, 'Cross Setback Toe depth found')
                    print(surveyToe, 'Cross Setback Toe depth found', file = f)
                
                # print(repr(cell)[1:16])
        
            # print(i, row)                        
                
    # print deep / shallow and summary of good/bad
        sheetSetBack = workbookPerf[workbookPerf.sheetnames[0]]
        surveyKB = int((sheetSurvey.cell(8,9).value[3:5]))
        surveyHeelGL = surveyHeel - surveyKB
        surveyToeGL = surveyToe - surveyKB
        print('\nSummary')
        print('\nSummary', file = f)
        print('\nDeepest perf        Shallowest perf')    
        print('\nDeepest perf        Shallowest perf',file=f)          
        print(' ', deepPerf,'             ', shallowPerf) 
        print(' ', deepPerf,'             ', shallowPerf,file=f)   
        print('\nToe Set-back        Heel set-back') 
        print('\nToe Set-back        Heel set-back',file=f) 
        print(' ',round(sheetSetBack[cellToe].value), '             ', round(sheetSetBack[cellHeel].value))  
        print(' ',round(sheetSetBack[cellToe].value), '             ', round(sheetSetBack[cellHeel].value),file=f) 
        print('\nSurvey Toe SB (GL)   Survey Heel SB (GL)')    
        print('\nSurvey Toe SB (GL)   Survey Heel SB (GL)',file=f)          
        print(' ', surveyToeGL,'             ', surveyHeelGL) 
        print(' ', surveyToeGL,'             ', surveyHeelGL,file=f)
        print()
        print(file = f)
        # Future -  need error handling made for this next part 
        # AF22 and AF26 cell values might change 
        if int(deepPerf) < int(sheetSetBack[cellToe].value):
            print('Toe perf is within toe set-back line, Toe perfs are good.')
            print('Toe perf is within toe set-back line, Toe perfs are good.',file=f) 
        if int(deepPerf) >= int(sheetSetBack[cellToe].value):
            print('ERROR, Toe perf is deeper than toe set-back.')
            print('ERROR, Toe perf is deeper than toe set-back.',file=f)
        if int(shallowPerf) > int(sheetSetBack[cellHeel].value):
            print('Heel perf is within heel set-back line, Heel perfs are good.')
            print('Heel perf is within heel set-back line, Heel perfs are good.',file=f)
        if int(shallowPerf) <= int(sheetSetBack[cellHeel].value):
            print('ERROR, Heel perf is shallower than heel set-back.')
            print('ERROR, Heel perf is shallower than heel set-back.',file=f)
        if int(surveyHeelGL) != int(sheetSetBack[cellHeel].value):
            print('Warning!, Survey heel setback does not equal perf sheet setback.', surveyHeelGL-int(sheetSetBack[cellHeel].value),'ft difference.')
            print('Warning!, Survey heel setback does not equal perf sheet setback.', surveyHeelGL-int(sheetSetBack[cellHeel].value),'ft difference.',file=f)
        if int(surveyToeGL) != int(sheetSetBack[cellToe].value):
            print('Warning!, Survey toe setback does not equal perf sheet toe setback.', surveyToeGL-int(sheetSetBack[cellToe].value),'ft difference.')
            print('Warning!, Survey toe setback does not equal perf sheet toe setback.', surveyToeGL-int(sheetSetBack[cellToe].value),'ft difference.',file=f)
        print('\n')
        
        
        # print(conflict)

