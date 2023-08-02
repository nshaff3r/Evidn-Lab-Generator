def lab_creator(lab_name, lab_file, start_date, end_date, building_name):
    import numpy as np
    import pandas as pd
    from time import strftime
    from constructor import Lab, Building
    from datetime import date, timedelta, datetime
       
   
    lab_name = lab_file.split(" ")[0][-4:]
    
    #------------------------------------------------------------------------------------------
    #CLEANING DATA
    
    data = pd.read_csv(lab_file,skiprows=[0])
    #Drop NaN and 0 values
    data = data.dropna(subset=['Value'])
    data = data.loc[data["Value"]!=0]
    
    data = data.reset_index()
    data = data.drop("index", axis = 1)
    
    date_format = "%m/%d/%Y"
    
    #-------------------------------------------------------------------------------------------
    #CALCULATING MONTH AVERAGE
    
    month_start_date = datetime(int(start_date.year), int(start_date.month), 1)
   # print(month_start_date)
  
    month_sum = 0
    month_count = 0
    month_index_list = [] #list to host indices of all relevant dates (from month) in the spreadsheet
    
    for i in np.arange(0, data.shape[0]):
        day = data["Date"][i]
        #get the month in the current value of the column
        day_split_list = day.split(' ')
        month_date_obj = datetime.strptime(day_split_list[0],date_format)
        if month_start_date <= month_date_obj <= end_date:
            month_sum = month_sum + data['Value'][i]
            month_count = month_count + 1
            
    month_average = month_sum/month_count
    
    #updating month_results_dict
    months_average = {1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 
                          7:0, 8:0, 9:0, 10:0, 11:0, 12:0}
    month_update = {int(start_date.month) : month_average}
    months_average.update(month_update)
    
    #----------------------------------------------------------------------------------------------
    #CALCULATING WEEK AVERAGE

    week_sum = 0
    week_count = 0

    for i in np.arange(0,data.shape[0]):
        initial_date = data["Date"][i]
        initial_date_list = initial_date.split(" ")
        date_obj = datetime.strptime(initial_date_list[0], date_format)
        if start_date <= date_obj <= end_date:
            week_sum = week_sum + data['Value'][i]
            week_count = week_count + 1

    week_average = week_sum/week_count
    baseline_average = float(input("What is " + lab_name +" in " + building_name +"'s baseline average? "))
    energy_saved = baseline_average - week_average
    miles = abs(energy_saved/3.9E-4)
    phones = abs(energy_saved/8.22E-6)
    homes = abs(energy_saved/7.93)
    alerts = int(input("How many alerts did " + lab_name + " recieve? "))
    lab = Lab(lab_name, baseline_average, week_average, months_average, energy_saved, 
                     miles, phones, homes, alerts, building_name)

    return lab
    

def building_average(lab_list, name):
    import numpy as np
    import pandas as pd
    from time import strftime
    from constructor import Lab, Building
    from datetime import date, timedelta, datetime
    
    #PROMPTING USER TO INPUT DATES (START OF WEEK AND END OF WEEK)
    
    start_year = int(input('Enter the start year (e.g. 2023): '))
    start_month = int(input('Enter the start month (a number e.g. May = 5): '))
    start_day = int(input('Enter the start day (e.g. 20): '))
    start_date = datetime(start_year, start_month, start_day)
    
    end_year = int(input('Enter the end year (e.g. 2023): '))
    end_month = int(input('Enter the end month (a number e.g. May = 5): '))
    end_day = int(input('Enter the end day (e.g. 26): '))
    end_date = datetime(end_year, end_month, end_day)
    
    building_name = name
    
    lab_week_sum = 0
    lab_week_count = 0
    lab_average = 0 #how do i actually set this
    week = 0 #how do i actually set this
    curr_building = Building(building_name, building_name, lab_average, week)
    
    response = "y"
    ignored_labs = []
    while response == "y":
        response = str(input("Do you have a lab that should not be included in the building average? Answer y/n: "))
        if response == "y":
            ignored_labs.append(str(input("Type in the lab number like L221: ")))
        else:
            print("No more labs to add")
    
    
    for i in np.arange(0, len(lab_list)):
        #lab_file = lab_list
        #when I have more than one lab file
        lab_file = lab_list[i]
        print(lab_file)
        lab_name = lab_file.split(" ")[0][-4:]
        lab = lab_creator(lab_name, lab_file, start_date, end_date, building_name)
        if lab_name not in ignored_labs:
            lab_week_sum = lab.week_avg + lab_week_sum
            lab_week_count += 1
        
        curr_building.add_lab(lab)

    curr_building.lab_average = lab_week_sum/lab_week_count
    return curr_building

def add_labs():
    import numpy as np
    import os
    lab_list = []
    currBuilding = str(input("What is the name of the building folder that contains your labs? "))
    directory = os.getcwd() + "/" + currBuilding
    if not os.path.exists(directory):
        print(f"Cannot find a folder with the name {currBuilding} in your current directory.\n"
              f"Check the spelling, and make sure it's at {os.getcwd()}/\n  ")
    else:
        for filename in os.listdir(directory):
            f = os.path.join(directory, filename)
            # checking if it is a file
            if os.path.isfile(f):
                lab_list.append(f)

    if len(lab_list) > 0:
        print("Your current lab files are: ")
        for i in np.arange(0,len(lab_list)):
            print(lab_list[i])
    
    #bool(lab_list) = True (not empty)
    if bool(lab_list):
        curr_building = building_average(lab_list, currBuilding)
        return curr_building
    else:
        print("You have no labs to process!")
        return