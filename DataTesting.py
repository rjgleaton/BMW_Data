from typing import List, Tuple, Set, Dict
import pandas as pd
import pdb as pdb
from datetime import datetime, timedelta
import datetime #one of the functions doesnt work unless I have this line, idk
import openpyxl, math

PATH: str = "MD04/MD04_10152021.xlsx"
saftey_stock: int
stock: int

def main():
    while(True):
        #Want to work with xlsx files but data is currently stored as .csv, run this code to convert
        #read_file = pd.read_csv(r"MD04/MD04 data until end of year test parts 10.15.21.csv")
        #read_file.to_excel("MD04/MD04_10152021.xlsx", index = None, header = True)

        wb = openpyxl.load_workbook(PATH)
        sheet = wb.active
        data = pd.DataFrame(sheet.values)

        material = input("What material do you want the health status for? ")
        date = input ("For what date (mm/dd/yyyy)? ")

        #Find material at specified date, loop over that material for the next
        #10 days and average out the values
        avg: List = []

        #Find last line index for that material and that date
        df_filtered = data[data[0] == material]

        #Get saftey stock
        ss_holder = df_filtered[df_filtered[12] == 2]
        #Work around because I'm not super familiar with pandas tbh
        for index, row in ss_holder.iterrows():
            if(row[6] == "SafeSt"):
                saftey_stock = abs(row[7])
                break
            #If not saftey stock value specified in file set to 0 by default
            else:
                saftey_stock = 0
                break


        
        #Sort data by only data on the date, get stock
        '''
        df_filtered_date = df_filtered[df_filtered[2] == date]
        if(df_filtered_date.empty):
            print("No data present for",date)
            return
        else:
            stock_holder = df_filtered_date[df_filtered_date[12]==df_filtered_date[12].max()]
            stock = stock_holder[8].values[0]
            health_score = get_health_score(stock, saftey_stock)
            avg.append(health_score)
            print("Health Status of",material,"on",date,"is",health_score,", and stock is",stock)
        '''
        mm, dd, yyyy = map(int, date.split('/'))
        date_obj = datetime.datetime(yyyy, mm, dd)

        for i in range(10):
            td = datetime.timedelta(days = i)
            new_date = date_obj + td
            formatted_date = datetime.date.strftime(new_date, "%m/%d/%Y")

            df_filtered_date = df_filtered[df_filtered[2] == formatted_date]
            if(df_filtered_date.empty):
                print("No data present for",formatted_date)
            else:
                stock_holder = df_filtered_date[df_filtered_date[12]==df_filtered_date[12].max()]
                stock = stock_holder[8].values[0]
                health_score = get_health_score(stock, saftey_stock)
                avg.append(health_score)
                print("Health Status of",material,"on",date,"is",health_score,", and stock is",stock)

        print("Average Health Status is",sum(avg)/len(avg))


def get_health_score(stock: int, saftey_stock: int) -> float:
    #Change value of k to affect attitude of curve
    return sigmoid(SS = (stock/saftey_stock), k = 0.3)

def sigmoid(SS: int, k: float) -> float:
    return (((1/(1+math.exp(-k*SS)))-(1/2))*2)*100


if __name__ == "__main__":
    main()
