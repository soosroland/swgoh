from csv import writer
import xlsxwriter
import pandas as pd

# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook('RiseOfEmpirePlatoon.xlsx')
PlanForPhase = 3
 
# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()

unit_current = ['TRIPLEZERO','50RT', 'AAYLASECURA','ADMIRALACKBAR','ADMIRALPIETT','ADMIRALRADDUS','AHSOKATANO','FULCRUMAHSOKA',
'AMILYNHOLDO','ARCTROOPER501ST','ASAJVENTRESS','AURRA_SING','B1BATTLEDROIDV2','B2SUPERBATTLEDROID','BARRISSOFFEE','BASTILASHAN',
'BASTILASHANDARK','BAZEMALBUS','BB8','BENSOLO','BIGGSDARKLIGHTER','BISTAN','BOKATAN','BOBAFETT','BOBAFETTSCION','BODHIROOK','BOSSK',
'BOUSHH','BT1','C3POLEGENDARY','CADBANE','CANDEROUSORDO','HOTHHAN','PHASMA','CARADUNE','CARTHONASI','CASSIANANDOR','CC2224',
'CHEWBACCALEGENDARY','CHIEFCHIRPA','CHIEFNEBIT','CHIRRUTIMWE','CHOPPERS3','CLONESERGEANTPHASEI','CLONEWARSCHEWBACCA','COLONELSTARCK',
'COMMANDERAHSOKA','COMMANDERLUKESKYWALKER','CORUSCANTUNDERWORLDPOLICE','COUNTDOOKU','CT210408','CT5555','CT7567','DARKTROOPER',
'DARTHMALAK','DARTHMALGUS','MAUL','DARTHNIHILUS','DARTHREVAN','DARTHSIDIOUS','DARTHSION','DARTHTALON','DARTHTRAYA','VADER','DASHRENDAR',
'DATHCHA','DEATHTROOPER','DENGAR','DIRECTORKRENNIC','DROIDEKA','BADBATCHECHO','EETHKOTH','EIGHTHBROTHER','EMBO','EMPERORPALPATINE',
'ENFYSNEST','EWOKELDER','EWOKSCOUT','EZRABRIDGERS3','FENNECSHAND','FIFTHBROTHER','FINN','FIRSTORDEREXECUTIONER','FIRSTORDEROFFICERMALE',
'FIRSTORDERSPECIALFORCESPILOT','FIRSTORDERTROOPER','FIRSTORDERTIEPILOT','GAMORREANGUARD','GARSAXON','ZEBS3','GRIEVOUS','GENERALHUX','GENERALKENOBI',
'GENERALSKYWALKER','VEERS','GEONOSIANBROODALPHA','GEONOSIANSOLDIER','GEONOSIANSPY','GRANDADMIRALTHRAWN','GRANDINQUISITOR','GRANDMASTERYODA',
'GRANDMOFFTARKIN','GREEDO','GREEFKARGA','HANSOLO','HERASYNDULLAS3','HERMITYODA','HK47','HONDO','HOTHREBELSCOUT','HOTHREBELSOLDIER',
'BADBATCHHUNTER','IDENVERSIOEMPIRE','MAGNAGUARD','IG11','IG86SENTINELDROID','IG88','IMAGUNDI','IMPERIALPROBEDROID','IMPERIALSUPERCOMMANDO',
'JABBATHEHUTT','JANGOFETT','JAWA','JAWAENGINEER','JAWASCAVENGER','JEDIKNIGHTCONSULAR','ANAKINKNIGHT','JEDIKNIGHTGUARDIAN','JEDIKNIGHTLUKE',
'JEDIKNIGHTREVAN','JEDIMASTERKENOBI','GRANDMASTERLUKE','JOLEEBINDO','JUHANI','JYNERSO','K2SO','KANANJARRUSS3','KIADIMUNDI','KITFISTO','KRRSANTAN',
'KUIIL','KYLEKATARN','KYLOREN','KYLORENUNMASKED','L3_37','ADMINISTRATORLANDO','LOBOT','LOGRAY','LORDVADER','LUKESKYWALKER','LUMINARAUNDULI',
'MACEWINDU','MAGMATROOPER','MARAJADE','MAULS7','MISSIONVAO','HUMANTHUG','MOFFGIDEONS1','MONMOTHMA','MOTHERTALZIN','NIGHTSISTERACOLYTE',
'NIGHTSISTERINITIATE','NIGHTSISTERSPIRIT','NIGHTSISTERZOMBIE','NINTHSISTER','NUTEGUNRAY','OLDBENKENOBI','DAKA','BADBATCHOMEGA','PADMEAMIDALA',
'PAO','PAPLOO','PLOKOON','POE','POGGLETHELESSER','PRINCESSLEIA','QIRA','QUIGONJINN','R2D2_LEGENDARY','RANGETROOPER','HOTHLEIA','EPIXFINN','EPIXPOE',
'RESISTANCEPILOT','RESISTANCETROOPER','GLREY','REYJEDITRAINING','REY','ROSETICO','ROYALGUARD','SABINEWRENS3','SANASTARROS','SAVAGEOPRESS',
'SCARIFREBEL','SECONDSISTER','SEVENTHSISTER','SHAAKTI','SHORETROOPER','SITHASSASSIN','SITHTROOPER','SITHPALPATINE','SITHMARAUDER','FOSITHTROOPER',
'UNDERCOVERLANDO','SNOWTROOPER','STARKILLER','STORMTROOPER','STORMTROOPERHAN','SUNFAC','SUPREMELEADERKYLOREN','T3_M4','TALIA','BADBATCHTECH','TEEBO',
'ARMORER','THEMANDALORIAN','THEMANDALORIANBESKARARMOR','THIRDSISTER','C3POCHEWBACCA','TIEFIGHTERPILOT','TUSKENRAIDER','TUSKENSHAMAN','UGNAUGHT',
'URORRURRR','YOUNGCHEWBACCA','SMUGGLERCHEWBACCA','SMUGGLERHAN','VISASMARR','WAMPA','WATTAMBOR','WEDGEANTILLES','WICKET', 'BADBATCHWRECKER','YOUNGHAN',
'YOUNGLANDO','ZAALBAR','ZAMWESELL']

unit_req_p1 = [4,0,0,0,1,0,0,1,
1,1,1,0,1,0,1,0,
0,0,6,2,0,0,0,1,5,0,1,
0,3,6,0,2,0,0,0,0,0,1,
5,0,0,0,0,1,0,1,
3,0,0,1,0,0,0,1,
0,2,3,0,10,2,3,0,10,3,0,
0,0,1,0,2,0,1,0,1,5,
0,0,0,0,0,1,0,0,2,
2,0,0,1,0,0,5,0,4,
10,0,2,0,0,4,0,4,
0,0,1,5,0,7,0,1,0,0,
1,0,0,0,1,2,1,2,1,
0,1,0,1,0,0,0,0,8,
4,0,0,2,0,0,0,0,3,0,0,
1,0,0,0,1,0,1,3,0,0,0,
0,0,1,0,0,0,0,0,0,0,
1,0,0,1,1,0,0,0,2,
0,1,0,1,0,0,0,0,4,0,1,0,0,
0,0,0,3,0,0,0,1,5,0,
0,0,0,0,0,1,0,0,0,0,
0,0,0,0,0,0,0,1,1,0,1,
1,0,7,0,3,1,0,1,3,
0,0,0,0,1,0,6,0,0,0,0,
0,1,0]

unit_req_p2 = [0,0,2,1,0,1,0,0,
0,1,1,0,0,2,0,1,
0,0,2,13,0,1,0,0,12,1,0,
1,1,2,1,0,0,0,0,0,0,0,
2,0,0,0,0,1,0,1,
11,1,0,1,1,1,0,0,
4,10,0,0,6,0,0,0,0,0,0,
0,2,1,0,1,0,0,0,2,4,
1,0,0,1,0,0,0,0,2,
0,0,0,0,0,0,1,1,1,
4,2,0,0,0,4,6,4,
0,0,2,1,0,1,2,0,0,0,
0,0,0,0,0,0,0,4,0,
0,0,1,0,0,0,0,0,3,
1,2,0,0,1,0,0,2,1,0,0,
1,0,0,2,2,1,0,1,1,0,0,
0,1,0,8,0,0,0,1,0,0,
0,0,0,0,0,0,0,1,1,
0,0,0,0,0,0,0,0,2,0,0,0,0,
0,0,3,2,0,1,0,0,1,0,
0,0,0,0,0,0,0,4,0,1,
0,1,4,0,0,0,3,0,0,0,1,
0,0,1,0,2,0,0,0,0,
0,1,1,0,0,0,4,1,0,0,1,
0,1,0]

unit_req_p3 = [1,1,0,0,0,0,0,0,
0,0,0,1,1,0,1,0,
2,0,3,7,0,0,0,0,10,1,0,
0,0,3,1,0,0,0,1,0,0,0,
0,0,0,1,0,1,0,0,
9,3,1,0,0,1,0,1,
2,14,1,0,1,0,0,0,1,0,0,
0,0,2,1,0,0,0,0,1,0,
0,1,1,0,0,1,0,0,0,
1,0,0,0,1,0,3,1,1,
1,0,0,0,1,1,0,2,
1,0,0,2,0,2,2,1,2,1,
1,0,0,0,0,1,0,2,0,
0,0,0,0,0,0,1,0,4,
2,5,7,0,0,0,1,0,3,0,0,
0,0,0,1,0,0,0,5,6,0,1,
1,0,1,11,0,0,0,0,1,1,
0,0,0,0,0,2,0,0,3,
0,0,0,0,0,0,0,1,4,0,2,1,1,
0,2,7,1,0,1,0,0,0,1,
0,0,1,0,0,1,0,7,0,1,
2,0,2,0,0,0,4,0,1,0,0,
0,0,5,0,0,0,2,0,0,
0,0,0,0,0,0,3,0,0,0,1,
0,0,0]

unit_req_p4 = [0,3,1,1,0,0,0,0,
0,0,3,2,0,1,0,0,
1,1,1,6,0,2,1,0,6,0,0,
0,1,0,1,0,0,0,0,0,0,0,
2,0,1,0,0,0,0,0,
7,0,1,0,2,0,0,1,
6,4,1,0,0,0,0,0,0,1,1,
1,0,1,1,0,0,1,1,1,0,
0,0,0,0,0,0,1,0,0,
0,1,1,1,0,0,0,0,3,
4,2,0,1,0,0,0,1,
0,0,0,2,0,5,1,2,0,0,
0,1,0,0,0,2,0,0,0,
0,0,0,0,0,1,0,0,4,
2,3,8,1,0,0,0,0,2,0,1,
0,1,1,0,0,0,0,2,7,0,0,
0,0,1,6,0,0,0,0,0,3,
1,1,1,0,0,0,0,0,2,
0,0,0,0,1,0,0,0,1,1,4,0,0,
0,0,7,1,0,0,0,0,1,1,
2,1,2,0,0,0,1,9,0,0,
0,2,13,0,0,0,7,0,0,1,0,
0,0,2,0,0,1,0,0,0,
0,0,0,1,0,1,0,1,1,0,0,
0,2,0]

unit_req_p5 = [0,3,0,0,0,0,0,1,
1,0,0,0,0,0,2,2,
0,2,0,2,0,1,0,1,9,2,0,
1,0,2,0,0,0,0,0,0,0,0,
2,0,0,1,1,1,0,1,
6,3,0,1,1,0,1,0,
9,7,0,0,0,1,0,0,0,0,0,
0,0,0,0,0,1,2,0,0,1,
0,0,0,0,0,0,0,1,1,
2,0,1,0,0,0,0,0,9,
10,0,0,1,0,0,10,5,
0,0,0,0,0,0,0,1,0,1,
1,1,1,0,0,3,1,0,0,
0,1,0,0,0,0,0,0,12,
3,0,8,1,0,0,0,0,0,0,0,
0,0,2,1,0,1,1,0,5,0,0,
1,1,0,4,0,0,1,1,0,0,
0,0,1,0,0,0,2,1,2,
0,0,0,0,0,0,0,0,2,0,0,1,0,
1,0,5,0,0,1,0,0,0,0,
0,0,0,0,0,1,2,8,0,0,
0,1,13,0,0,0,4,1,2,2,0,
1,0,0,0,0,0,0,2,0,
0,0,1,0,0,1,0,3,1,0,0,
0,1,0]

unit_req_p6 = [0,0,0,0,0,1,0,0,
0,0,2,1,2,1,0,6,
0,0,1,10,1,0,0,0,9,0,0,
0,0,0,0,0,0,0,0,0,0,0,
1,0,1,0,0,1,0,0,
9,0,1,2,0,0,0,1,
0,11,1,0,0,0,0,1,3,0,5,
0,0,1,0,0,0,0,0,0,0,
0,0,0,0,0,1,0,0,1,
0,0,0,0,0,0,0,1,5,
0,0,2,1,1,0,0,1,
1,0,0,6,0,5,2,3,0,0,
0,0,2,0,1,2,2,2,2,
0,0,1,1,1,1,0,0,0,
1,7,9,0,0,0,0,1,3,0,1,
0,0,0,1,1,0,1,0,10,0,0,
0,0,2,8,0,0,0,0,0,2,
1,1,0,2,1,0,1,0,1,
1,0,0,1,0,0,0,0,0,1,4,0,0,
1,0,11,0,2,0,1,0,2,1,
0,0,0,0,1,0,1,3,0,2,
2,0,0,0,0,0,8,0,0,1,0,
0,0,1,0,0,0,1,0,2,
0,0,1,0,0,0,2,0,0,0,0,
1,1,1]


file_name = "E:\\Roland\\swgoh\\csvs\\unit-export.csv"

# write header
worksheet.write(0, 0, "Character name")
worksheet.write(0, 1, "R5 current")
worksheet.write(0, 2, "P1 req")
worksheet.write(0, 3, "P1 status")
worksheet.write(0, 4, "R6 current")
worksheet.write(0, 5, "P2 req")
worksheet.write(0, 6, "P2 status")
worksheet.write(0, 7, "R7 current")
worksheet.write(0, 8, "P3 req")
worksheet.write(0, 9, "P3 status")
worksheet.write(0,10, "R8 current")
worksheet.write(0,11, "P4 req")
worksheet.write(0,12, "P4 status")
worksheet.write(0,13, "R9 current")
worksheet.write(0,14, "P5 req")
worksheet.write(0,15, "P5 status")
worksheet.write(0,16, "P6 req")
worksheet.write(0,17, "P6 status")

for i in range(len(unit_current)):
    #18. Ben Solo 23. Boba Scion 54. Darth Malgus 119. Jabba 147. LV 153. Maul 184. Rey 190. Sana 199. SEE 204. SK
    if i != 19 and i != 24 and i != 55 and i != 120 and i != 148 and i != 154 and i != 185 and i != 191 and i != 200 and i != 205:
        # reading the csv file
        df = pd.read_csv(file_name, encoding="ISO-8859-1", on_bad_lines='skip', sep=',')

        column = 0
        # write character name
        worksheet.write(i+1, column, str(unit_current[i]))
        column = column + 1 

        # write R5 data
        worksheet.write(i+1, column, str(df[df['Relic Tier'] >= 5]['Relic Tier'].count()))
        column = column + 1 

        # write P1 required
        worksheet.write(i+1, column, unit_req_p1[i])
        column = column + 1 

        # write if we have enough for P1
        if ( df[df['Relic Tier'] >= 5]['Relic Tier'].count() >= unit_req_p1[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 

        # write R6 data
        worksheet.write(i+1, column, str(df[df['Relic Tier'] >= 6]['Relic Tier'].count()))
        column = column + 1 

        # write P2 required
        worksheet.write(i+1, column, unit_req_p2[i])
        column = column + 1 

        # write if we have enough for P2
        if ( df[df['Relic Tier'] >= 6]['Relic Tier'].count() >= unit_req_p2[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 

        # write R7 data
        worksheet.write(i+1, column, str(df[df['Relic Tier'] >= 7]['Relic Tier'].count()))
        column = column + 1 

        # write P3 required
        worksheet.write(i+1, column, unit_req_p3[i])
        column = column + 1 

        # write if we have enough for P3
        if ( df[df['Relic Tier'] >= 7]['Relic Tier'].count() >= unit_req_p3[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 
        
        # write R8 data
        worksheet.write(i+1, column, str(df[df['Relic Tier'] >= 8]['Relic Tier'].count()))
        column = column + 1 

        # write P4 required
        worksheet.write(i+1, column, unit_req_p4[i])
        column = column + 1 

        # write if we have enough for P4
        if ( df[df['Relic Tier'] >= 8]['Relic Tier'].count() >= unit_req_p4[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 
        
        # write R9 data
        worksheet.write(i+1, column, str(df[df['Relic Tier'] >= 9]['Relic Tier'].count()))
        column = column + 1 

        # write P5 required
        worksheet.write(i+1, column, unit_req_p5[i])
        column = column + 1 

        # write if we have enough for P5
        if ( df[df['Relic Tier'] >= 9]['Relic Tier'].count() >= unit_req_p5[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 
        
        # write P6 required
        worksheet.write(i+1, column, unit_req_p6[i])
        column = column + 1 

        # write if we have enough for P6
        if ( df[df['Relic Tier'] >= 9]['Relic Tier'].count() >= unit_req_p6[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 
        

        # Phase 6 highest
        if PlanForPhase == 6:
            if unit_req_p6[i] >= max(unit_req_p5[i],unit_req_p4[i],unit_req_p3[i],unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, "R5..9 needed: " + str(unit_req_p6[i]))
            elif unit_req_p5[i] >= max(unit_req_p6[i],unit_req_p4[i],unit_req_p3[i],unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, "R5..9 needed: " + str(unit_req_p5[i]))
            elif unit_req_p4[i] >= max(unit_req_p3[i],unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, "R9 needed: " + str(max(unit_req_p6[i],unit_req_p5[i])) + " R8 needed: " + str(unit_req_p4[i]))
            elif unit_req_p3[i] >= max(unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, "R9 needed: " + str(max(unit_req_p6[i],unit_req_p5[i])) + " R7 needed: " + str(unit_req_p3[i]))
            elif unit_req_p2[i] >= unit_req_p1[i]:
                worksheet.write(i+1, column, "R9 needed: " + str(max(unit_req_p6[i],unit_req_p5[i])) + " R6 needed: " + str(unit_req_p2[i]))
            else:
                worksheet.write(i+1, column, "R9 needed: " + str(max(unit_req_p6[i],unit_req_p5[i])) + " R8 needed: " + str(unit_req_p4[i])
                + " R7 needed: " + str(unit_req_p3[i]) + " R6 needed: " + str(unit_req_p2[i]) + " R5 needed: " + str(unit_req_p1[i]))
        elif PlanForPhase == 3:
            if unit_req_p1[i] == 0 and unit_req_p2[i] == 0 and unit_req_p3[i] == 0:
                worksheet.write(i+1, column, "")
            elif unit_req_p3[i] >= max(unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, str(unit_current[i]) + " R7 needed: " + str(unit_req_p3[i]))
            elif unit_req_p2[i] >= unit_req_p1[i]:
                worksheet.write(i+1, column, str(unit_current[i]) + " R7 needed: " + str(unit_req_p3[i]) + " R6 needed: " + str(unit_req_p2[i]))
            else:
                worksheet.write(i+1, column, str(unit_current[i]) + " R7 needed: " + str(unit_req_p3[i]) + " R6 needed: " + str(unit_req_p2[i]) + " R5 needed: " + str(unit_req_p1[i]))

        # print(str(unit_current[i]) + 
        # " R5: " + str(df[df['Relic Tier'] >= 5]['Relic Tier'].count()) + 
        # " R6: " + str(df[df['Relic Tier'] >= 6]['Relic Tier'].count()) + 
        # " R7: " + str(df[df['Relic Tier'] >= 7]['Relic Tier'].count()) + 
        # " R8: " + str(df[df['Relic Tier'] >= 8]['Relic Tier'].count()) + 
        # " R9: " + str(df[df['Relic Tier'] >= 9]['Relic Tier'].count()))
    elif i == 19 or i == 24 or i == 55 or i == 120 or i == 148 or i == 154 or i == 185 or i == 191 or i == 200 or i == 205:
        # reading the csv file
        df = pd.read_csv(file_name, encoding="ISO-8859-1", on_bad_lines='skip', sep=',')
        column = 0
        # write character name
        worksheet.write(i+1, column, str(unit_current[i]))
        column = column + 1 

        # write R5 data
        worksheet.write(i+1, column, str(df[df['Gear Tier'] >= 5]['Gear Tier'].count()))
        column = column + 1 

        # write P1 required
        worksheet.write(i+1, column, unit_req_p1[i])
        column = column + 1 

        # write if we have enough for P1
        if ( df[df['Gear Tier'] >= 5]['Gear Tier'].count() >= unit_req_p1[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 

        # write R6 data
        worksheet.write(i+1, column, str(df[df['Gear Tier'] >= 6]['Gear Tier'].count()))
        column = column + 1 

        # write P2 required
        worksheet.write(i+1, column, unit_req_p2[i])
        column = column + 1 

        # write if we have enough for P2
        if ( df[df['Gear Tier'] >= 6]['Gear Tier'].count() >= unit_req_p2[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 

        # write R7 data
        worksheet.write(i+1, column, str(df[df['Gear Tier'] >= 7]['Gear Tier'].count()))
        column = column + 1 

        # write P3 required
        worksheet.write(i+1, column, unit_req_p3[i])
        column = column + 1 

        # write if we have enough for P3
        if ( df[df['Gear Tier'] >= 7]['Gear Tier'].count() >= unit_req_p3[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 
        
        # write R8 data
        worksheet.write(i+1, column, str(df[df['Gear Tier'] >= 8]['Gear Tier'].count()))
        column = column + 1 

        # write P4 required
        worksheet.write(i+1, column, unit_req_p4[i])
        column = column + 1 

        # write if we have enough for P4
        if ( df[df['Gear Tier'] >= 8]['Gear Tier'].count() >= unit_req_p4[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 
        
        # write R9 data
        worksheet.write(i+1, column, str(df[df['Gear Tier'] >= 9]['Gear Tier'].count()))
        column = column + 1 

        # write P5 required
        worksheet.write(i+1, column, unit_req_p5[i])
        column = column + 1 

        # write if we have enough for P5
        if ( df[df['Gear Tier'] >= 9]['Gear Tier'].count() >= unit_req_p5[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 
        
        # write P6 required
        worksheet.write(i+1, column, unit_req_p6[i])
        column = column + 1 

        # write if we have enough for P6
        if ( df[df['Gear Tier'] >= 9]['Gear Tier'].count() >= unit_req_p6[i] ):
            worksheet.write(i+1, column, "")
        else:
            worksheet.write(i+1, column, "!!!!!!")
        column = column + 1 

        # Phase 6 highest
        if PlanForPhase == 6:
            if unit_req_p6[i] >= max(unit_req_p5[i],unit_req_p4[i],unit_req_p3[i],unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, "R5..9 needed: " + str(unit_req_p6[i]))
            elif unit_req_p5[i] >= max(unit_req_p6[i],unit_req_p4[i],unit_req_p3[i],unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, "R5..9 needed: " + str(unit_req_p5[i]))
            elif unit_req_p4[i] >= max(unit_req_p3[i],unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, "R9 needed: " + str(max(unit_req_p6[i],unit_req_p5[i])) + " R8 needed: " + str(unit_req_p4[i]))
            elif unit_req_p3[i] >= max(unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, "R9 needed: " + str(max(unit_req_p6[i],unit_req_p5[i])) + " R7 needed: " + str(unit_req_p3[i]))
            elif unit_req_p2[i] >= unit_req_p1[i]:
                worksheet.write(i+1, column, "R9 needed: " + str(max(unit_req_p6[i],unit_req_p5[i])) + " R6 needed: " + str(unit_req_p2[i]))
            else:
                worksheet.write(i+1, column, "R9 needed: " + str(max(unit_req_p6[i],unit_req_p5[i])) + " R8 needed: " + str(unit_req_p4[i])
                + " R7 needed: " + str(unit_req_p3[i]) + " R6 needed: " + str(unit_req_p2[i]) + " R5 needed: " + str(unit_req_p1[i]))
        elif PlanForPhase == 3:
            if unit_req_p1[i] == 0 and unit_req_p2[i] == 0 and unit_req_p3[i] == 0:
                worksheet.write(i+1, column, "")
            elif unit_req_p3[i] >= max(unit_req_p2[i],unit_req_p1[i]):
                worksheet.write(i+1, column, str(unit_current[i]) + " R7 needed: " + str(unit_req_p3[i]))
            elif unit_req_p2[i] >= unit_req_p1[i]:
                worksheet.write(i+1, column, str(unit_current[i]) + " R7 needed: " + str(unit_req_p3[i]) + " R6 needed: " + str(unit_req_p2[i]))
            else:
                worksheet.write(i+1, column, str(unit_current[i]) + " R7 needed: " + str(unit_req_p3[i]) + " R6 needed: " + str(unit_req_p2[i]) + " R5 needed: " + str(unit_req_p1[i]))

    file_name = "E:\\Roland\\swgoh\\csvs\\unit-export(" + str(i+1) + ").csv"
# close the file
workbook.close()