"""
guide
1 creat ini file named 'config.ini' and put in same folder python program
2 creat state file named 'ProgramState.xlsx' and put in same folder python program
3 check path of 1.mt5 folder 2.mt5 program 
4 modifi [TesterInputs] in 'config.ini'
5 creat folder name report(name in config file of step 1) in mt5 folder


"""

import subprocess
from numpy import NaN, true_divide
import numpy as np
import pandas as pd
import os
import pathlib
import configparser
import math
import shutil
import xml.etree.ElementTree as ET

programEnd = False


ProgramPath = pathlib.Path(__file__).parent.resolve()
State_Path = os.path.join(ProgramPath , 'ProgramState.xlsx')

# get symbol and time in ListSymbol file: creat data for main loop
State_data = pd.DataFrame ( pd.read_excel(State_Path) )

# get config.ini file path
Config_path = os.path.join(ProgramPath , 'config.ini')
config = configparser.ConfigParser()
config.read(Config_path)

# get mt5 path
mt5_folder_path = 'C:\\Users\\admin\AppData\Roaming\MetaQuotes\Terminal\A75E3666C6C81392587C985DC08C4FC4'
mt5_program_path = 'C:\\Program Files\\MetaTrader 5 backtest\\terminal64.exe'
set_file_path = os.path.join(mt5_folder_path , 'Profiles\Tester\Ea-ZicZac.set')

for index , symbol in enumerate(State_data.symbol ):
    print(index , symbol)
    # 1 optimize 70% data
    if State_data.state[index] != "ok":
        #  Optimize Config file    
        config['Tester']['Symbol'] = symbol+"_Tickstory"
        config['Tester']['Model'] = "1" # 1 — "1 minute OHLC", 4 — "Every tick based on real ticks"
        config['Tester']['optimization'] = "1" # 0 — optimization disabled, 1 — "Slow complete algorithm"
        config['Tester']['optimizationcriterion'] = "6" # 0 — the maximum balance, 6 — a custom optimization
        config['Tester']['FromDate'] = str( State_data.startY[index] ) + ".01.01"
        config['Tester']['ToDate'] = str ( State_data.endY[index])+ ".01.01"
        filename = "opt_" + symbol +"_"+ str( State_data.startY[index] ) +"_"+ str ( State_data.endY[index])
        config['Tester']['Report'] ='\\' + os.path.join("reports" , symbol , filename)
        with open(Config_path, 'w') as configw: config.write(configw) # save to config file
        # run mt5
        subprocess.call([mt5_program_path, '/config:'+Config_path])
        optReport_path = os.path.join(mt5_folder_path, "reports", symbol, filename+ ".xml")
        target_path = os.path.join(mt5_folder_path, "reports", symbol, filename+ ".xlsx")
        input_parameter =[]
        if os.path.exists(optReport_path) :
            print('final optimize ' + symbol)
            # get best parameter from xml file AND save xml file to xlsx
            input_data = ET.parse(optReport_path)
            root = input_data.getroot()
            _prefix = '{urn:schemas-microsoft-com:office:spreadsheet}'
            row = 0
            for child in root.iter(_prefix+'Row') : row +=1
            data =[]
            for child in root.iter(_prefix +'Data') :
                data.append(child.text)
            collume = int(len(data) / row)
            # save xml to xlsx
            num = np.array(data)
            reshaped = num.reshape(row,collume)
            df = pd.DataFrame(reshaped)
            df.to_excel(target_path, header=False , index= False)
            # get best para
            if data[collume+7] != "0" :
                input_parameter = data[collume+10 : collume*2]
    else : continue
            
    # 2 backtest 70% onTick of best Value in optimize
    if len(input_parameter)>0 :
        print(input_parameter)
        i=0
        for x in config['TesterInputs']:
            sp = config['TesterInputs'][x].split("||")
            if sp[4] == "Y" : 
                sp[0] = input_parameter[i]
                i+=1
                config['TesterInputs'][x] = sp[0] + "||" + sp[1] + "||" + sp[2] + "||" + sp[3] + "||" + sp[4]
        config['Tester']['Model'] = "4" # model every tick base real tick
        config['Tester']['optimization'] = "0"
        filename_bt = "bt_" + symbol +"_"+ str( State_data.startY[index] ) +"_"+ str ( State_data.endY[index])
        config['Tester']['Report'] ='\\' + os.path.join("reports" ,symbol, filename_bt)
        with open(Config_path, 'w') as configw: config.write(configw) # save to config file

        subprocess.call([mt5_program_path, '/config:'+Config_path])
        bt_path = os.path.join(mt5_folder_path ,"reports" ,symbol, filename_bt+ ".htm")
    else : 
        State_data.state[index] = "ok" 
        State_data.to_excel(State_Path , index= False) 
        continue
    
    if os.path.exists(bt_path) :
        config['Tester']['FromDate'] = str ( State_data.endY[index])+ ".01.01"
        config['Tester']['ToDate'] = "2022.01.01"
        filename_fw = "fw_" + symbol +"_"+ str( State_data.endY[index] ) +"_"+ "2022"
        config['Tester']['Report'] ='\\' + os.path.join("reports" ,symbol, filename_fw)
        with open(Config_path, 'w') as configw: config.write(configw) # save to config file
        subprocess.call([mt5_program_path, '/config:'+Config_path])
        fw_path = os.path.join(mt5_folder_path ,"reports" ,symbol, filename_fw+ ".htm")
    
    if os.path.exists(fw_path) :
        config['Tester']['FromDate'] = "2022.01.01"
        config['Tester']['ToDate'] = "2022.06.01"
        filename_vt = "vt_" + symbol
        config['Tester']['Report'] ='\\' + os.path.join("reports" ,symbol, filename_vt)
        with open(Config_path, 'w') as configw: config.write(configw) # save to config file
        subprocess.call([mt5_program_path, '/config:'+Config_path])
        vt_path = os.path.join(mt5_folder_path ,"reports" ,symbol, filename_vt+ ".htm")

     # save to state
    if os.path.exists(vt_path) :
        State_data.state[index] = "ok" 
        State_data.to_excel(State_Path , index= False)      

    # run main program in loop each symbol
        # 1 optimize 70% data 
        # 2 backtest 70% onTick of best Value in optimize
        # 3 forward Test 30% data to 01.01.2022
        # 4 Validate test 01.01.2022 to 01.06.2022
    
        

#subprocess.call(['C:\\Users\\admin\\Documents\\MT5Portable\\MetaTrader 5 backtest\\terminal64.exe', '/config:config1.ini'])

    
