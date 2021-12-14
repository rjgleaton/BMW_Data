from typing import List, Tuple, Set, Dict
from xlwt import Workbook
import pandas as pd
import pdb as pdb
import openpyxl, math

def main():

    #Justing using 4 files for beginning stage testing, want to keep it simple for now
    path_list: list = ["130_Report/130_Report_1005.xlsx", "130_Report/130_Report_1006.xlsx","130_Report/130_Report_1008.xlsx","130_Report/130_Report_1011.xlsx"]

    #For each file in the list, create a dictionary where keys are material ID (column 4 or E) data is list containing DoS, Saftey Stock (its ratio compared to stock), Supplier Location, change in DoS (initially 0), health status
    #For each file update all items, note change in DoS, calculate and update health status, print each time

    parts: Dict[str, list] = dict()
    count = 0
    for path in path_list:
        count += 1
        #Reading Excel file, using a work around because pd.read_excel was being a pain
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        data = pd.DataFrame(sheet.values)
        for index, row in data.iterrows():
            #First row is just headers
            if index == 0:
                continue

            label = row[4]
            if row[31] != 0:
                SS = row[32]/row[31]
            else: 
                SS = 0

            #TODO actually calculate how many days passed... currently just does a false version making use of the known excel sheet names
            if label in parts:
                change_in_days = int(path[22:26])-int(path_list[count-2][22:26]) 
                change_DoS = (parts[label][0] - row[20])/change_in_days
                if change_DoS < 0:
                    change_DoS = 0
            else: 
                change_DoS = 0
            #DoS, SS, Supplier Location, change in DoS, health status
            parts[label] = [row[20], SS, row[37], change_DoS, calc_health(row[20], SS, row[37], change_DoS)]

        #pdb.set_trace()
        writebook = Workbook()
        sheet1 = writebook.add_sheet('Sheet 1')
        #Write to column 49
        for index, key in enumerate(parts):
            #pdb.set_trace()
            sheet1.write(index+1, 0, parts[key][4])
        writebook.save("test%s.xls" % count) 
    
    pdb.set_trace()
    

def calc_health(DoS: int, SS: int, supplier_loc: str, change_DoS: float) -> float:
    #Base Case
    if DoS <= 0:
        return 0
    
    #All values chosen as 'optimal' and the amount by which a value can affect the score is more or less arbitrarily chosen at this point
    initial_score: float = 30.0
    
    #We consider one week of supply to be optimal to begin with
    if(DoS <= 7):
        initial_score = initial_score*(DoS/7)

    #Check for saftey stock
    initial_score = SS_consideration(SS, supplier_loc, initial_score)

    #Check to see if DoS is changing more than 1 per day
    #Again just throw this into a sigmoid to make it scale as the value gets larger, k is set to -1 so it should act as a standard sigmoid function
    #TODO probably should warn parts handler if this is the case aswell
    if change_DoS > 1.0:
        initial_score = initial_score-(0.2*sigmoid(change_DoS, k = -1))

    return max(0.0, initial_score)


def SS_consideration(SS: int, supplier_loc: str, initial_score: float) -> float:
    #We consider a ration of SS to stock of at least 5 to be optimal
    #Base case for SS
    if SS == 0:
        #Check if supplier is close by
        #TODO make this function actually consider distance to supplier, not just be a 'is it in america?' check
        if supplier_loc == "United States":
            initial_score = initial_score * 0.7
        else: 
            initial_score = initial_score * 0.5
    #For now we just calculate this using a modified sigmoid function I stole from here: https://math.stackexchange.com/questions/1832177/sigmoid-function-with-fixed-bounds-and-variable-steepness-partially-solved
    elif SS >= 5:
        #Check if supplier is close by
        #TODO make this function actually consider distance to supplier, not just be a 'is it in america?' check
        if supplier_loc == "United States":
            initial_score = initial_score * (1-(0.3*sigmoid(SS, k = -0.1)))
        else:
            initial_score = initial_score * (1-(0.5*sigmoid(SS, k = -0.1)))
    
    return initial_score

def sigmoid(SS: int, k: float) -> float:
    return 1 - (1/(1+math.exp(-k*SS)))


if __name__ == "__main__":
    main()
