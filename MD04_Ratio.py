from typing import List, Tuple, Set, Dict
import pandas as pd
import pdb as pdb
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

        #Find last line index for that material and that date
        df_filtered = data[data[0] == material]

        #Get saftey stock
        ss_holder = df_filtered[df_filtered[12] == 2]
        #Work around because I'm not super familiar with pandas tbh
        for index, row in ss_holder.iterrows():
            if(row[6] == "SafeSt"):
                saftey_stock = abs(row[7])
                break
            else:
                saftey_stock = 0
                break


        
        #Sort data by only data on the date, get stock
        df_filtered_date = df_filtered[df_filtered[2] == date]
        if(df_filtered_date.empty):
            print("No data present for given date")
        else:
            stock_holder = df_filtered_date[df_filtered_date[12]==df_filtered_date[12].max()]
            stock = stock_holder[8].values[0]

        print("Health Status of",material,"is",get_health_score(stock, saftey_stock))




def get_health_score(stock: int, saftey_stock: int) -> float:
    #Change value of k to affect attitude of curve
    return sigmoid(SS = (stock/saftey_stock), k = 0.3)

def sigmoid(SS: int, k: float) -> float:
    return (((1/(1+math.exp(-k*SS)))-(1/2))*2)*100


if __name__ == "__main__":
    main()