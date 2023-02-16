import gzip
import tkinter
from tkinter import filedialog
import os
import pygsheets
import re
from colorama import Fore
import sys
import time

#the following likely will be for the temporary code
from datetime import datetime
import calendar

ascii_logo =  """      

     @@@@@@                                                                     
     @@@@@@                                                                     
    @@@@@@       @@@@@     @@@@  @@@@@     @@@@@   @@@@@@@@@   @@@@@@      @@@@ 
    @@@@@@       @@@@@     @@@@    @@@@@  @@@@@    @@@@@@@@@   @@@@@@@@    @@@@ 
   @@@@@@        @@@@@     @@@@     ,@@@@@@@@      @@@@        @@@@@@@@@   @@@@ 
   @@@@@@        @@@@@     @@@@       @@@@@@       @@@@@@@@@   @@@@  @@@@  @@@@ 
  @@@@@@         @@@@@     @@@@     &@@@@@@@@      @@@@        @@@@   @@@@ @@@@ 
  @@@@@@@@@@@@@   @@@@@@@@@@@@     @@@@@  @@@@@    @@@@@@@@@   @@@@    @@@@@@@@ 
 @@@@@@@@@@@@@      @@@@@@@@     @@@@@     @@@@@.  @@@@@@@@@   @@@@      @@@@@@ 

"""   
print(Fore.LIGHTMAGENTA_EX + ascii_logo)
print("===================Luxen's SMPE Top 10 Naton Log Data Scanner===================")
print("")
while True:
    try:
        print("Launching File Explorer...")
        print("")
        print("Please Select a folder containing the chat logs.")
        window = tkinter.Tk()
        window.withdraw()
        window.wm_attributes('-topmost', 1)
        logs_folder = filedialog.askdirectory()
        window.destroy()
        logs_list = os.listdir(logs_folder)
    except FileNotFoundError:
        print(Fore.WHITE + "Error: Log folder not given")
        askrelaunch = ""
        while askrelaunch.lower() != "y":
            askrelaunch = input("Relaunch File Explorer? (y/n)").lower()
            if askrelaunch == "n":    
                print("Exiting Program in 3 seconds.....")
                time.sleep(3)    
                sys.exit()
            elif askrelaunch != "y":
                print("Error: Invalid response: use 'y' for yes or 'n' for no")
        continue

    top_10_nation_list = {}
    for gzlog in logs_list:
        if gzlog.endswith('.gz') and (int(gzlog[0:4]) > 2021 or (int(gzlog[0:4]) == 2021 and int(gzlog[5:7]) >= 7)):
            print(Fore.BLUE + "Scanning {}".format(gzlog))
            log = gzip.open(logs_folder + "/" +gzlog, 'rt').read()
            log = log.replace('§b', "")
            log = log.replace("§8", "")
            last_nation_inst = log.rfind('.[ Nations ].')
            if last_nation_inst != -1:
                list_nation_end = log.find('<<<',last_nation_inst)
                checknationtype = log[log.find('Nation Name - ',last_nation_inst)+14:log.find('Nation Name - ',last_nation_inst)+33] == "Number of Residents"
                list_nation_page_num = log.find('<<<',last_nation_inst)+11
                #print(Fore.LIGHTCYAN_EX + str(log[last_nation_inst:list_nation_end]))
                if int(log[list_nation_page_num]) == 1 and checknationtype:
                    print(Fore.GREEN + "Found Top 10 Nations List")
                    nation_list_formatted = str(log[last_nation_inst:list_nation_end])
                    nation_list_formatted = nation_list_formatted.replace("[Render thread/INFO]: [System] [CHAT]", "")
                    nation_list_formatted = nation_list_formatted.replace("[ Nations ].__________________.oOo.", "")
                    nation_list_formatted = nation_list_formatted.replace("[main/INFO]: [CHAT]", "")
                    nation_list_formatted = nation_list_formatted.replace("[Render thread/INFO]: [CHAT]", "")

                    timestamp_list = list(re.finditer("(\[[0-9][0-9]\:[0-9][0-9]\:[0-9][0-9]\])", nation_list_formatted))
                    nations_res_count = {}
                    for i, timestamp in enumerate(timestamp_list):
                        if i != len(timestamp_list)-1:
                            next_timestamp = timestamp_list[i+1]
                            nation_info_inst = nation_list_formatted[timestamp.end(0)+ 1:next_timestamp.start(0)]
                            if "Nation Name - Number of Residents\n" in nation_info_inst:
                                continue
                            res_count = list(re.finditer("- \([0-9]+\)", nation_info_inst))[-1]
                            nations_res_count[nation_info_inst[1:res_count.start(0)-1]] = res_count.group(0)[3:-1]
                            

                    top_10_nation_list[gzlog[:-7]] = nations_res_count

    #temp
    for date, nation_info in top_10_nation_list.items():
        format_date = datetime.strptime(date[:10], "%Y-%m-%d")
        print(Fore.YELLOW + "---- Date: {} {}, {} ----".format(format_date.day, calendar.month_name[format_date.month], format_date.year))
        print("Log # for date: {}".format(date[11:]))
        iter_count = 1
        for nation, res_count in nation_info.items():
            print(Fore.GREEN +"#{} : ".format(iter_count) + Fore.LIGHTBLUE_EX + "Nation: " + Fore.LIGHTCYAN_EX + str(nation) + Fore.GREEN + " |" +
                  Fore.LIGHTBLUE_EX + " # of Residents: " + Fore.LIGHTCYAN_EX + str(res_count))
            iter_count += 1
        print("")


    #print(Fore.WHITE + str(top_10_nation_list))



    asktorepeat = ""
    while asktorepeat.lower() != "y":
        asktorepeat = input(Fore.LIGHTMAGENTA_EX + "Logs have been scanned. Open another folder? (y/n)")
        asktorepeat.lower()
        if asktorepeat == "n":    
            print("Exiting Program in 3 seconds.....")
            time.sleep(3)    
            sys.exit()
        elif asktorepeat != "y":
            print("Error: Invalid response: use 'y' for yes or 'n' for no")
    continue

        
