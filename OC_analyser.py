import xlwings as xw
import pandas as pd
import numpy as np 
import time
from datetime import datetime
def timee():
    timexx = datetime.now().strftime("%I:%M %p")
    return timexx
def spot_prob_N50(sheet):
    SN = []
    VAL = []
    SPOT = sheet.range('J3').expand().options(pd.DataFrame, index=False).value
    SN = SPOT['S.No'].tolist()
    VAL = SPOT['Value'].tolist()
    for a,b in zip(SN,VAL):
        if(b != 0):
            sheet.range('B4').value = ('=INDEX(NIFTY[underlyingValue],{})'.format(int(a)))
            break
def spot_prob_BNF(sheet):
    SN = []
    VAL = []
    SPOT = sheet.range('U3').expand().options(pd.DataFrame, index=False).value
    SN = SPOT['S.No'].tolist()
    VAL = SPOT['Value'].tolist()
    for a,b in zip(SN,VAL):
        if(b != 0):
            sheet.range('P4').value = ('=INDEX(BANKNIFTY[underlyingValue],{})'.format(int(a)))
            break
def spot_prob_FIN(sheet):
    SN = []
    VAL = []
    SPOT = sheet.range('AL3').expand().options(pd.DataFrame, index=False).value
    SN = SPOT['S.No'].tolist()
    VAL = SPOT['Value'].tolist()
    for a,b in zip(SN,VAL):
        if(b != 0):
            sheet.range('AD4').value = ('=INDEX(FINNIFTY[underlyingValue],{})'.format(int(a)))
            break
def N50_DATA(df,t):
    N_SP_ITM1 = [];N_SP_ITM2 = [];N_SP_ITM3 = [];N_SP_ITM4 = [];N_SP_ITM5 = [];N_SP_ITM6 = [];N_SP_ITM7 = [];N_SP_ITM8 = [];N_SP_ITM9 = []
    N_SP_ITM10 = [];N_SP_ITM11 = [];N_SP_ITM12 = [];N_SP_ATM = [];N_SP_OTM1 = [];N_SP_OTM2 = [];N_SP_OTM3 = [];N_SP_OTM4 = [];N_SP_OTM5 = []
    N_SP_OTM6 = [];N_SP_OTM7 = [];N_SP_OTM8 = [];N_SP_OTM9 = [];N_SP_OTM10 = [];N_SP_OTM11 = [];N_SP_OTM12 = []
    N_CE_OI = df['CE_OI'].tolist();N_CE_OI_CHG  = df['CE_OI_CHG'].tolist();N_CE_VOL = df['CE_VOL'].tolist();N_CE_IV = df['CE_IV'].tolist()
    N_CE_LTP_CHG = df['CE_LTP_CHG'].tolist();N_CE_LTP = df['CE_LTP'].tolist();N_STRIKE = df['STRIKE'].tolist()
    N_PE_OI = df['PE_OI'].tolist();N_PE_OI_CHG  = df['PE_OI_CHG'].tolist();N_PE_VOL = df['PE_VOL'].tolist();N_PE_IV = df['PE_IV'].tolist()
    N_PE_LTP_CHG = df['PE_LTP_CHG'].tolist();N_PE_LTP = df['PE_LTP'].tolist()
    out_file = xw.Book('Option_Chain_Analysis.xlsx')
    out_sht = out_file.sheets['N50_DATA']
    N_SP = np.array(out_sht.range('B2').options(ndim = 1).value) 
    try:
        loc = N_STRIKE.index(N_SP)
    except ValueError:
        print("Strike Price not found, change the value in the Excel Analysis:Data(sheet)\n")
        print('N50_SP:{}'.format(N_SP))
        exit()
    ATM_SP = N_STRIKE[24]
    shift = N_STRIKE[loc] - ATM_SP
    N_SP = ATM_SP + shift
    timexx = t
    for a,b,c,d,e,f,x,g,h,i,j,k,l in zip(N_CE_OI,N_CE_OI_CHG,N_CE_VOL,N_CE_IV,N_CE_LTP_CHG,N_CE_LTP,N_STRIKE,N_PE_LTP,N_PE_LTP_CHG,N_PE_IV,N_PE_VOL,N_PE_OI_CHG,N_PE_OI):
        if(x == (N_SP-600)):
            N_SP_ITM12.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-550)):
            N_SP_ITM11.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-500)):
            N_SP_ITM10.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-450)):
            N_SP_ITM9.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-400)):
            N_SP_ITM8.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-350)):
            N_SP_ITM7.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-300)):
            N_SP_ITM6.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-250)):
            N_SP_ITM5.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-200)):
            N_SP_ITM4.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-150)):
            N_SP_ITM3.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-100)):
            N_SP_ITM2.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP-50)):
            N_SP_ITM1.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == N_SP):
            N_SP_ATM.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+50)):
            N_SP_OTM1.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+100)):
            N_SP_OTM2.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+150)):
            N_SP_OTM3.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+200)):
            N_SP_OTM4.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+250)):
            N_SP_OTM5.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+300)):
            N_SP_OTM6.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+350)):
            N_SP_OTM7.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+400)):
            N_SP_OTM8.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+450)):
            N_SP_OTM9.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+500)):
            N_SP_OTM10.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+550)):
            N_SP_OTM11.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (N_SP+600)):
            N_SP_OTM12.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
    next_row_ITM12 = out_sht.range('BZ313').end('down').row + 1
    next_row_ITM11 = out_sht.range('BK313').end('down').row + 1
    next_row_ITM10 = out_sht.range('AV313').end('down').row + 1
    next_row_ITM9 = out_sht.range('AG313').end('down').row + 1
    next_row_ITM8 = out_sht.range('R313').end('down').row + 1
    next_row_ITM7 = out_sht.range('C313').end('down').row + 1
    next_row_ITM6 = out_sht.range('BZ210').end('down').row + 1
    next_row_ITM5 = out_sht.range('BK210').end('down').row + 1
    next_row_ITM4 = out_sht.range('AV210').end('down').row + 1
    next_row_ITM3 = out_sht.range('AG210').end('down').row + 1
    next_row_ITM2 = out_sht.range('R210').end('down').row + 1
    next_row_ITM1 = out_sht.range('C210').end('down').row + 1
    next_row_ATM = out_sht.range('C3').end('down').row + 1
    next_row_OTM1 = out_sht.range('R3').end('down').row + 1
    next_row_OTM2 = out_sht.range('AG3').end('down').row + 1
    next_row_OTM3 = out_sht.range('AV3').end('down').row + 1
    next_row_OTM4 = out_sht.range('BK3').end('down').row + 1
    next_row_OTM5 = out_sht.range('BZ3').end('down').row + 1
    next_row_OTM6 = out_sht.range('CO3').end('down').row + 1
    next_row_OTM7 = out_sht.range('C107').end('down').row + 1
    next_row_OTM8 = out_sht.range('R107').end('down').row + 1
    next_row_OTM9 = out_sht.range('AG107').end('down').row + 1
    next_row_OTM10 = out_sht.range('AV107').end('down').row + 1
    next_row_OTM11 = out_sht.range('BK107').end('down').row + 1
    next_row_OTM12 = out_sht.range('BZ107').end('down').row + 1
    out_sht.range('BZ' + str(next_row_ITM12)).value = N_SP_ITM12
    out_sht.range('BK' + str(next_row_ITM11)).value = N_SP_ITM11
    out_sht.range('AV' + str(next_row_ITM10)).value = N_SP_ITM10
    out_sht.range('AG' + str(next_row_ITM9)).value = N_SP_ITM9
    out_sht.range('R' + str(next_row_ITM8)).value = N_SP_ITM8
    out_sht.range('C' + str(next_row_ITM7)).value = N_SP_ITM7
    out_sht.range('BZ' + str(next_row_ITM6)).value = N_SP_ITM6
    out_sht.range('BK' + str(next_row_ITM5)).value = N_SP_ITM5
    out_sht.range('AV' + str(next_row_ITM4)).value = N_SP_ITM4
    out_sht.range('AG' + str(next_row_ITM3)).value = N_SP_ITM3
    out_sht.range('R' + str(next_row_ITM2)).value = N_SP_ITM2
    out_sht.range('C' + str(next_row_ITM1)).value = N_SP_ITM1
    out_sht.range('C' + str(next_row_ATM)).value = N_SP_ATM
    out_sht.range('R' + str(next_row_OTM1)).value = N_SP_OTM1
    out_sht.range('AG' + str(next_row_OTM2)).value = N_SP_OTM2
    out_sht.range('AV' + str(next_row_OTM3)).value = N_SP_OTM3
    out_sht.range('BK' + str(next_row_OTM4)).value = N_SP_OTM4
    out_sht.range('BZ' + str(next_row_OTM5)).value = N_SP_OTM5
    out_sht.range('CO' + str(next_row_OTM6)).value = N_SP_OTM6
    out_sht.range('C' + str(next_row_OTM7)).value = N_SP_OTM7
    out_sht.range('R' + str(next_row_OTM8)).value = N_SP_OTM8
    out_sht.range('AG' + str(next_row_OTM9)).value = N_SP_OTM9
    out_sht.range('AV' + str(next_row_OTM10)).value = N_SP_OTM10
    out_sht.range('BK' + str(next_row_OTM11)).value = N_SP_OTM11
    out_sht.range('BZ' + str(next_row_OTM12)).value = N_SP_OTM12
def BNF_DATA(df,t):
    B_SP_ITM1 = [];B_SP_ITM2 = [];B_SP_ITM3 = [];B_SP_ITM4 = [];B_SP_ITM5 = [];B_SP_ITM6 = [];B_SP_ITM7 = [];B_SP_ITM8 = [];B_SP_ITM9 = []
    B_SP_ITM10 = [];B_SP_ITM11 = [];B_SP_ITM12 = [];B_SP_ATM = [];B_SP_OTM1 = [];B_SP_OTM2 = [];B_SP_OTM3 = [];B_SP_OTM4 = [];B_SP_OTM5 = []
    B_SP_OTM6 = [];B_SP_OTM7 = [];B_SP_OTM8 = [];B_SP_OTM9 = [];B_SP_OTM10 = [];B_SP_OTM11 = [];B_SP_OTM12 = []
    B_CE_OI = df['CE_OI'].tolist();B_CE_OI_CHG = df['CE_OI_CHG'].tolist();B_CE_VOL = df['CE_VOL'].tolist();B_CE_IV = df['CE_IV'].tolist()
    B_CE_LTP_CHG = df['CE_LTP_CHG'].tolist();B_CE_LTP = df['CE_LTP'].tolist();B_STRIKE = df['STRIKE'].tolist()
    B_PE_OI = df['PE_OI'].tolist();B_PE_OI_CHG = df['PE_OI_CHG'].tolist();B_PE_VOL = df['PE_VOL'].tolist();B_PE_IV = df['PE_IV'].tolist()
    B_PE_LTP_CHG = df['PE_LTP_CHG'].tolist();B_PE_LTP = df['PE_LTP'].tolist() 
    out_file = xw.Book('Option_Chain_Analysis.xlsx')
    out_sht = out_file.sheets['BNF_DATA']
    B_SP = np.array(out_sht.range('B2').options(ndim = 1).value) 
    try:
        loc = B_STRIKE.index(B_SP)
    except ValueError:
        print("Strike Price not found, change the value in the Excel Analysis:Data(sheet)\n")
        print('BNF_SP:{}'.format(B_SP))
        exit()
    ATM_SP = B_STRIKE[24]
    shift = B_STRIKE[loc] - ATM_SP
    B_SP = ATM_SP + shift
    timexx = t
    for a,b,c,d,e,f,x,g,h,i,j,k,l in zip(B_CE_OI,B_CE_OI_CHG,B_CE_VOL,B_CE_IV,B_CE_LTP_CHG,B_CE_LTP,B_STRIKE,B_PE_LTP,B_PE_LTP_CHG,B_PE_IV,B_PE_VOL,B_PE_OI_CHG,B_PE_OI):
        if(x == (B_SP-1200)):
            B_SP_ITM12.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-1100)):
            B_SP_ITM11.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-1000)):
            B_SP_ITM10.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-900)):
            B_SP_ITM9.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-800)):
            B_SP_ITM8.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-700)):
            B_SP_ITM7.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-600)):
            B_SP_ITM6.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-500)):
            B_SP_ITM5.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-400)):
            B_SP_ITM4.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-300)):
            B_SP_ITM3.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-200)):
            B_SP_ITM2.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP-100)):
            B_SP_ITM1.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == B_SP):
            B_SP_ATM.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+100)):
            B_SP_OTM1.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+200)):
            B_SP_OTM2.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+300)):
            B_SP_OTM3.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+400)):
            B_SP_OTM4.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+500)):
            B_SP_OTM5.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+600)):
            B_SP_OTM6.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+700)):
            B_SP_OTM7.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+800)):
            B_SP_OTM8.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+900)):
            B_SP_OTM9.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+1000)):
            B_SP_OTM10.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+1100)):
            B_SP_OTM11.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (B_SP+1200)):
            B_SP_OTM12.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
    next_row_ITM12 = out_sht.range('BZ313').end('down').row + 1
    next_row_ITM11 = out_sht.range('BK313').end('down').row + 1
    next_row_ITM10 = out_sht.range('AV313').end('down').row + 1
    next_row_ITM9 = out_sht.range('AG313').end('down').row + 1
    next_row_ITM8 = out_sht.range('R313').end('down').row + 1
    next_row_ITM7 = out_sht.range('C313').end('down').row + 1
    next_row_ITM6 = out_sht.range('BZ210').end('down').row + 1
    next_row_ITM5 = out_sht.range('BK210').end('down').row + 1
    next_row_ITM4 = out_sht.range('AV210').end('down').row + 1
    next_row_ITM3 = out_sht.range('AG210').end('down').row + 1
    next_row_ITM2 = out_sht.range('R210').end('down').row + 1
    next_row_ITM1 = out_sht.range('C210').end('down').row + 1
    next_row_ATM = out_sht.range('C3').end('down').row + 1
    next_row_OTM1 = out_sht.range('R3').end('down').row + 1
    next_row_OTM2 = out_sht.range('AG3').end('down').row + 1
    next_row_OTM3 = out_sht.range('AV3').end('down').row + 1
    next_row_OTM4 = out_sht.range('BK3').end('down').row + 1
    next_row_OTM5 = out_sht.range('BZ3').end('down').row + 1
    next_row_OTM6 = out_sht.range('CO3').end('down').row + 1
    next_row_OTM7 = out_sht.range('C107').end('down').row + 1
    next_row_OTM8 = out_sht.range('R107').end('down').row + 1
    next_row_OTM9 = out_sht.range('AG107').end('down').row + 1
    next_row_OTM10 = out_sht.range('AV107').end('down').row + 1
    next_row_OTM11 = out_sht.range('BK107').end('down').row + 1
    next_row_OTM12 = out_sht.range('BZ107').end('down').row + 1
    out_sht.range('BZ' + str(next_row_ITM12)).value = B_SP_ITM12
    out_sht.range('BK' + str(next_row_ITM11)).value = B_SP_ITM11
    out_sht.range('AV' + str(next_row_ITM10)).value = B_SP_ITM10
    out_sht.range('AG' + str(next_row_ITM9)).value = B_SP_ITM9
    out_sht.range('R' + str(next_row_ITM8)).value = B_SP_ITM8
    out_sht.range('C' + str(next_row_ITM7)).value = B_SP_ITM7
    out_sht.range('BZ' + str(next_row_ITM6)).value = B_SP_ITM6
    out_sht.range('BK' + str(next_row_ITM5)).value = B_SP_ITM5
    out_sht.range('AV' + str(next_row_ITM4)).value = B_SP_ITM4
    out_sht.range('AG' + str(next_row_ITM3)).value = B_SP_ITM3
    out_sht.range('R' + str(next_row_ITM2)).value = B_SP_ITM2
    out_sht.range('C' + str(next_row_ITM1)).value = B_SP_ITM1
    out_sht.range('C' + str(next_row_ATM)).value = B_SP_ATM
    out_sht.range('R' + str(next_row_OTM1)).value = B_SP_OTM1
    out_sht.range('AG' + str(next_row_OTM2)).value = B_SP_OTM2
    out_sht.range('AV' + str(next_row_OTM3)).value = B_SP_OTM3
    out_sht.range('BK' + str(next_row_OTM4)).value = B_SP_OTM4
    out_sht.range('BZ' + str(next_row_OTM5)).value = B_SP_OTM5
    out_sht.range('CO' + str(next_row_OTM6)).value = B_SP_OTM6
    out_sht.range('C' + str(next_row_OTM7)).value = B_SP_OTM7
    out_sht.range('R' + str(next_row_OTM8)).value = B_SP_OTM8
    out_sht.range('AG' + str(next_row_OTM9)).value = B_SP_OTM9
    out_sht.range('AV' + str(next_row_OTM10)).value = B_SP_OTM10
    out_sht.range('BK' + str(next_row_OTM11)).value = B_SP_OTM11
    out_sht.range('BZ' + str(next_row_OTM12)).value = B_SP_OTM12 
def FIN_DATA(df,t):
    F_SP_ITM1 = [];F_SP_ITM2 = [];F_SP_ITM3 = [];F_SP_ITM4 = [];F_SP_ITM5 = [];F_SP_ITM6 = [];F_SP_ITM7 = [];F_SP_ITM8 = [];F_SP_ITM9 = []
    F_SP_ITM10 = [];F_SP_ITM11 = [];F_SP_ITM12 = [];F_SP_ATM = [];F_SP_OTM1 = [];F_SP_OTM2 = [];F_SP_OTM3 = [];F_SP_OTM4 = [];F_SP_OTM5 = []
    F_SP_OTM6 = [];F_SP_OTM7 = [];F_SP_OTM8 = [];F_SP_OTM9 = [];F_SP_OTM10 = [];F_SP_OTM11 = [];F_SP_OTM12 = []
    F_CE_OI = df['CE_OI'].tolist();F_CE_OI_CHG  = df['CE_OI_CHG'].tolist();F_CE_VOL = df['CE_VOL'].tolist();F_CE_IV = df['CE_IV'].tolist()
    F_CE_LTP_CHG = df['CE_LTP_CHG'].tolist();F_CE_LTP = df['CE_LTP'].tolist();F_STRIKE = df['STRIKE'].tolist()
    F_PE_OI = df['PE_OI'].tolist();F_PE_OI_CHG  = df['PE_OI_CHG'].tolist();F_PE_VOL = df['PE_VOL'].tolist();F_PE_IV = df['PE_IV'].tolist()
    F_PE_LTP_CHG = df['PE_LTP_CHG'].tolist();F_PE_LTP = df['PE_LTP'].tolist()
    out_file = xw.Book('Option_Chain_Analysis.xlsx')
    out_sht = out_file.sheets['FIN_DATA']
    F_SP = np.array(out_sht.range('B2').options(ndim = 1).value) 
    try:
        loc = F_STRIKE.index(F_SP)
    except ValueError:
        print("Strike Price not found, change the value in the Excel Analysis:Data(sheet)\n")
        print('FIN_SP:{}'.format(F_SP))
        print(F_STRIKE)
        exit()
    ATM_SP = F_STRIKE[24]
    shift = F_STRIKE[loc] - ATM_SP
    F_SP = ATM_SP + shift
    timexx = t
    for a,b,c,d,e,f,x,g,h,i,j,k,l in zip(F_CE_OI,F_CE_OI_CHG,F_CE_VOL,F_CE_IV,F_CE_LTP_CHG,F_CE_LTP,F_STRIKE,F_PE_LTP,F_PE_LTP_CHG,F_PE_IV,F_PE_VOL,F_PE_OI_CHG,F_PE_OI):
        if(x == (F_SP-600)):
            F_SP_ITM12.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-550)):
            F_SP_ITM11.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-500)):
            F_SP_ITM10.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-450)):
            F_SP_ITM9.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-400)):
            F_SP_ITM8.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-350)):
            F_SP_ITM7.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-300)):
            F_SP_ITM6.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-250)):
            F_SP_ITM5.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-200)):
            F_SP_ITM4.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-150)):
            F_SP_ITM3.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-100)):
            F_SP_ITM2.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP-50)):
            F_SP_ITM1.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == F_SP):
            F_SP_ATM.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+50)):
            F_SP_OTM1.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+100)):
            F_SP_OTM2.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+150)):
            F_SP_OTM3.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+200)):
            F_SP_OTM4.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+250)):
            F_SP_OTM5.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+300)):
            F_SP_OTM6.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+350)):
            F_SP_OTM7.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+400)):
            F_SP_OTM8.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+450)):
            F_SP_OTM9.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+500)):
            F_SP_OTM10.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+550)):
            F_SP_OTM11.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
        elif(x == (F_SP+600)):
            F_SP_OTM12.append([a,b,c,d,e,f,x,g,h,i,j,k,l,timexx])
    next_row_ITM12 = out_sht.range('BZ313').end('down').row + 1
    next_row_ITM11 = out_sht.range('BK313').end('down').row + 1
    next_row_ITM10 = out_sht.range('AV313').end('down').row + 1
    next_row_ITM9 = out_sht.range('AG313').end('down').row + 1
    next_row_ITM8 = out_sht.range('R313').end('down').row + 1
    next_row_ITM7 = out_sht.range('C313').end('down').row + 1
    next_row_ITM6 = out_sht.range('BZ210').end('down').row + 1
    next_row_ITM5 = out_sht.range('BK210').end('down').row + 1
    next_row_ITM4 = out_sht.range('AV210').end('down').row + 1
    next_row_ITM3 = out_sht.range('AG210').end('down').row + 1
    next_row_ITM2 = out_sht.range('R210').end('down').row + 1
    next_row_ITM1 = out_sht.range('C210').end('down').row + 1
    next_row_ATM = out_sht.range('C3').end('down').row + 1
    next_row_OTM1 = out_sht.range('R3').end('down').row + 1
    next_row_OTM2 = out_sht.range('AG3').end('down').row + 1
    next_row_OTM3 = out_sht.range('AV3').end('down').row + 1
    next_row_OTM4 = out_sht.range('BK3').end('down').row + 1
    next_row_OTM5 = out_sht.range('BZ3').end('down').row + 1
    next_row_OTM6 = out_sht.range('CO3').end('down').row + 1
    next_row_OTM7 = out_sht.range('C107').end('down').row + 1
    next_row_OTM8 = out_sht.range('R107').end('down').row + 1
    next_row_OTM9 = out_sht.range('AG107').end('down').row + 1
    next_row_OTM10 = out_sht.range('AV107').end('down').row + 1
    next_row_OTM11 = out_sht.range('BK107').end('down').row + 1
    next_row_OTM12 = out_sht.range('BZ107').end('down').row + 1
    out_sht.range('BZ' + str(next_row_ITM12)).value = F_SP_ITM12
    out_sht.range('BK' + str(next_row_ITM11)).value = F_SP_ITM11
    out_sht.range('AV' + str(next_row_ITM10)).value = F_SP_ITM10
    out_sht.range('AG' + str(next_row_ITM9)).value = F_SP_ITM9
    out_sht.range('R' + str(next_row_ITM8)).value = F_SP_ITM8
    out_sht.range('C' + str(next_row_ITM7)).value = F_SP_ITM7
    out_sht.range('BZ' + str(next_row_ITM6)).value = F_SP_ITM6
    out_sht.range('BK' + str(next_row_ITM5)).value = F_SP_ITM5
    out_sht.range('AV' + str(next_row_ITM4)).value = F_SP_ITM4
    out_sht.range('AG' + str(next_row_ITM3)).value = F_SP_ITM3
    out_sht.range('R' + str(next_row_ITM2)).value = F_SP_ITM2
    out_sht.range('C' + str(next_row_ITM1)).value = F_SP_ITM1
    out_sht.range('C' + str(next_row_ATM)).value = F_SP_ATM
    out_sht.range('R' + str(next_row_OTM1)).value = F_SP_OTM1
    out_sht.range('AG' + str(next_row_OTM2)).value = F_SP_OTM2
    out_sht.range('AV' + str(next_row_OTM3)).value = F_SP_OTM3
    out_sht.range('BK' + str(next_row_OTM4)).value = F_SP_OTM4
    out_sht.range('BZ' + str(next_row_OTM5)).value = F_SP_OTM5
    out_sht.range('CO' + str(next_row_OTM6)).value = F_SP_OTM6
    out_sht.range('C' + str(next_row_OTM7)).value = F_SP_OTM7
    out_sht.range('R' + str(next_row_OTM8)).value = F_SP_OTM8
    out_sht.range('AG' + str(next_row_OTM9)).value = F_SP_OTM9
    out_sht.range('AV' + str(next_row_OTM10)).value = F_SP_OTM10
    out_sht.range('BK' + str(next_row_OTM11)).value = F_SP_OTM11
    out_sht.range('BZ' + str(next_row_OTM12)).value = F_SP_OTM12
while True:
    start = time.time()
    print('----------------------')
    ex_file = xw.Book('Option_Chain_Original.xlsx')
    sht = ex_file.sheets['Analysis']
    ex_file.api.RefreshAll()
    time.sleep(3)
    ex_file.api.RefreshAll()
    time.sleep(4)
    ti = timee()
    check_N50 = sht.range('B4').value
    check_BNF = sht.range('P4').value
    check_FIN = sht.range('AD4').value
    if(check_N50 == 0):
        spot_prob_N50(sht)
    if(check_BNF == 0):
        spot_prob_BNF(sht)
    if(check_FIN == 0):
        spot_prob_FIN(sht)
    df_N50 = sht.range('B18').expand().options(pd.DataFrame, index=False).value
    df_BNF = sht.range('P18').expand().options(pd.DataFrame, index=False).value
    #df_FIN = sht.range('AD18').expand().options(pd.DataFrame, index=False).value
    N50_DATA(df_N50,ti)
    BNF_DATA(df_BNF,ti)
    #FIN_DATA(df_FIN,ti)
    end = time.time()
    print("time taken:",(end-start),"Sec")
    print('----------------------')
    time.sleep(165)
