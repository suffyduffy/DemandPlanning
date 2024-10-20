# -*- coding: utf-8 -*-
"""
"" SIT DSC 2302 (2023) lab05 Inventory Performance Simulaton Tool (Simulation algo)
"""

# Step 0:
# Create the Main GUI Window --------------------------------------------------

import tkinter as tk
import random
import time

from tkinter import TclError, ttk, Text
from tkinter import *
from tkinter import filedialog as fd

from tkinter.messagebox import showerror, showwarning, showinfo
import pandas as pd
from pandas.tseries.offsets import DateOffset
import numpy as np

from datetime import date, datetime
from dateutil.relativedelta import relativedelta

# main Window
MainWindow = tk.Tk()
MainWindow.title('Inventory Performance Simulaiton')
MainWindow.geometry("910x450+50+50")
MainWindow.resizable(True, True)

# Global variable with default value = 4 weeks
TotalSKU = 20
df1 = pd.DataFrame()
dfPolicy = pd.DataFrame()

ROPVar = tk.StringVar()
OrderQtyVar = tk.StringVar()
Record_values = pd.Series()

# ----------------

# Step 1:
# Create NoteBook  ------------------------------------------------------------

MyNotebook = ttk.Notebook(MainWindow)
MyNotebook.pack(pady=10, expand=True)

# create frames
FrameForecast = ttk.Frame(MyNotebook, width=900, height=400)
FrameLT = ttk.Frame(MyNotebook, width=900, height=400)
FrameTarget = ttk.Frame(MyNotebook, width=900, height=400)
FrameSim = ttk.Frame(MyNotebook, width=700, height=400)

FrameForecast.pack(fill='both', expand=True)
FrameLT.pack(fill='both', expand=True)
FrameTarget.pack(fill='both', expand=True)
FrameSim.pack(fill=tk.Y, expand=True, anchor=tk.E)

# add frames to notebook

MyNotebook.add(FrameForecast, text='Demand Forecast')
MyNotebook.add(FrameLT, text='Lead Time & Review Period')
MyNotebook.add(FrameTarget, text='Performance Target')
MyNotebook.add(FrameSim, text="Performance Evaluation")

# Step 2:
# Create treeview for forecast (input) ----------------------------------------

ColumnForecast = ('SerialNo', 'SKU Code', 'SKU Name', "Month 1", "Month 2", "Month 3", "Month 4", \
                  "Month 5", "Month 6", "Month 7", "Month 8", "Month 9", "Month 10", \
                  "Month 11", "Month 12")
ViewForecast = ttk.Treeview(FrameForecast, columns=ColumnForecast, height=15, show='headings')

ViewForecast.column("SerialNo", width=50)
ViewForecast.column("SKU Code", width=100)
ViewForecast.column("SKU Name", width=100)
ViewForecast.column("Month 1", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 2", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 3", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 4", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 5", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 6", width=90, anchor=tk.CENTER)
ViewForecast.column("Month 7", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 8", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 9", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 10", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 11", width=80, anchor=tk.CENTER)
ViewForecast.column("Month 12", width=80, anchor=tk.CENTER)

ViewForecast.heading('SerialNo', text='No')
ViewForecast.heading('SKU Code', text='SKU Code')
ViewForecast.heading('SKU Name', text='SKU Name')
ViewForecast.heading('Month 1', text='Month 1')
ViewForecast.heading('Month 2', text='Month 2')
ViewForecast.heading('Month 3', text='Month 3')
ViewForecast.heading('Month 4', text='Month 4')
ViewForecast.heading('Month 5', text='Month 5')
ViewForecast.heading('Month 6', text='Month 6')
ViewForecast.heading('Month 7', text='Month 7')
ViewForecast.heading('Month 8', text='Month 8')
ViewForecast.heading('Month 9', text='Month 9')
ViewForecast.heading('Month 10', text='Month 10')
ViewForecast.heading('Month 11', text='Month 11')
ViewForecast.heading('Month 12', text='Month 12')

for i in range(TotalSKU):
    ViewForecast.insert('', tk.END, iid=i + 1, text="", values=(i + 1, "", "", "", "", "", "", "", \
                                                                "", "", "", "", "", "", ""))

ViewForecast.pack(ipadx=20, ipady=20, anchor=tk.CENTER, fill=tk.BOTH, expand=True)

# Add a scroll bar
ForecastScrollbar = ttk.Scrollbar(
    ViewForecast,
    orient='horizontal',
    command=ViewForecast.xview
)
ViewForecast.configure(xscroll=ForecastScrollbar.set)
ForecastScrollbar.pack(ipadx=5, ipady=5, fill=tk.X, side=tk.BOTTOM)

# Step 3:
# Create treeview for supply lead time (input) --------------------------------

ColumnLT = ('SerialNo', 'SKU Code', 'SKU Name', "Min LT", "Mode LT", "Max LT", "Review Period")
ViewLT = ttk.Treeview(FrameLT, columns=ColumnLT, height=15, show='headings')

ViewLT.column("SerialNo", width=50)
ViewLT.column("SKU Code", width=100)
ViewLT.column("SKU Name", width=100)
ViewLT.column("Min LT", width=80, anchor=tk.CENTER)
ViewLT.column("Mode LT", width=80, anchor=tk.CENTER)
ViewLT.column("Max LT", width=80, anchor=tk.CENTER)
ViewLT.column("Review Period", width=80, anchor=tk.CENTER)

ViewLT.heading('SerialNo', text='No')
ViewLT.heading('SKU Code', text='SKU Code')
ViewLT.heading('SKU Name', text='SKU Name')
ViewLT.heading('Min LT', text='Min LT')
ViewLT.heading('Mode LT', text='Mode LT')
ViewLT.heading('Max LT', text='Max LT')
ViewLT.heading('Review Period', text='Review Period')

for i in range(TotalSKU):
    ViewLT.insert('', tk.END, iid=i + 1, text="", values=(i + 1, "", "", "", "", "", "", ""))

ViewLT.pack(ipadx=20, ipady=20, anchor=tk.CENTER, fill=tk.X, expand=True)

# Step 4:
# Create treeview for performance target ---------------------------------------

ColumnTarget = (
'SerialNo', 'SKU Code', 'SKU Name', "Target Yearly Turns", "Target Fill Rate", "Target Inventory Position")

ViewTarget = ttk.Treeview(FrameTarget, columns=ColumnTarget, height=15, show='headings')

ViewTarget.column("SerialNo", width=50)
ViewTarget.column("SKU Code", width=100)
ViewTarget.column("SKU Name", width=100)
ViewTarget.column("Target Yearly Turns", width=100, anchor=tk.CENTER)
ViewTarget.column("Target Fill Rate", width=100, anchor=tk.CENTER)
ViewTarget.column("Target Inventory Position", width=100, anchor=tk.CENTER)

ViewTarget.heading('SerialNo', text='No')
ViewTarget.heading('SKU Code', text='SKU Code')
ViewTarget.heading('SKU Name', text='SKU Name')
ViewTarget.heading('Target Yearly Turns', text='Target Yearly Turns')
ViewTarget.heading('Target Fill Rate', text='Target Fill Rate')
ViewTarget.heading('Target Inventory Position', text='Target Inventory Position')

for i in range(TotalSKU):
    ViewTarget.insert('', tk.END, iid=i + 1, text="", values=(i + 1, "", "", "", "", "", ""))

ViewTarget.pack(ipadx=20, ipady=20, anchor=tk.CENTER, fill=tk.BOTH, expand=True)

# Modify policy for what-if analysis


# Step 5:
# Create treeview for simulation result ---------------------------------------


ColumnSim = ('No', 'SKU Code', 'SKU Name', "ROP", "Order Qty", "Predicted Turns", "Predicted Fill Rate", \
             "Predicted Inventory Position", "Predicted On Hand", "Predicted Open Order")

ViewSim = ttk.Treeview(FrameSim, columns=ColumnSim, height=15, show='headings')

ViewSim.column("No", width=50)
ViewSim.column("SKU Code", width=100)
ViewSim.column("SKU Name", width=100)
ViewSim.column("ROP", width=80, anchor=tk.CENTER)
ViewSim.column("Order Qty", width=80, anchor=tk.CENTER)
ViewSim.column("Predicted Turns", width=80, anchor=tk.CENTER)
ViewSim.column("Predicted Fill Rate", width=80, anchor=tk.CENTER)
ViewSim.column("Predicted Inventory Position", width=80, anchor=tk.CENTER)
ViewSim.column("Predicted On Hand", width=80, anchor=tk.CENTER)
ViewSim.column("Predicted Open Order", width=80, anchor=tk.CENTER)

ViewSim.heading('No', text='No')
ViewSim.heading('SKU Code', text='SKU Code')
ViewSim.heading('SKU Name', text='SKU Name')
ViewSim.heading('ROP', text='ROP')
ViewSim.heading('Order Qty', text='Order Qty')
ViewSim.heading('Predicted Turns', text='Predicted Turns')
ViewSim.heading('Predicted Fill Rate', text='Predicted Fill Rate')
ViewSim.heading('Predicted Inventory Position', text='Predicted Inventory Position')
ViewSim.heading('Predicted On Hand', text='Predicted On Hand')
ViewSim.heading('Predicted Open Order', text='Predicted Open Order')

for i in range(TotalSKU):
    ViewSim.insert('', tk.END, iid=i + 1, text="", values=(i + 1, "", "", "", "", "", "", "", "", "", ""))

ViewSim.pack(ipadx=20, ipady=20, anchor=tk.E, fill=tk.BOTH, expand=True)


# Modify policy for selected item
def ChangePolicy():
    global Record_values
    global ROPVar
    global OrderQtyVar

    CurrIndex = Record_values[0]
    SKUCode = Record_values[1]
    SKUName = Record_values[2]

    # retrieve new policy from Entry boxes
    ROP = ROPVar.get()
    OrderQty = OrderQtyVar.get()

    # update the new policy
    ViewSim.item(CurrIndex, text=SKUCode, values=(CurrIndex, SKUCode, SKUName, ROP, OrderQty, "", "", "", "", ""))
    showinfo(title="policy", message="Policy has been updated!")


# Record selected item and show its policy to the entry boxes
def ResultItem_selected(event):
    global ROPVar
    global OrderQtyVar
    global Record_values

    try:
        Selected_Record_No = ViewSim.selection()
        Record = ViewSim.item(Selected_Record_No)
    except Exception as e:
        showinfo(title="Error", message=e)
    else:
        Record_values = Record['values']
        ROP = Record_values[3]
        OrderQty = Record_values[4]

        ROPEntry.delete(0, 'end')
        OrderQtyEntry.delete(0, 'end')
        ROPEntry.insert(0, str(ROP))
        OrderQtyEntry.insert(0, str(OrderQty))


ViewSim.bind('<<TreeviewSelect>>', ResultItem_selected)

# add two labels and two Entry boxes
ttk.Label(FrameSim, text="New ROP:").pack(padx=5, ipadx=2, ipady=2, side=tk.LEFT, expand=False)
ROPEntry = ttk.Entry(FrameSim, textvariable=ROPVar)
ROPEntry.pack(side=tk.LEFT, expand=False)

ttk.Label(FrameSim, text="New Order Qty:").pack(padx=10, ipadx=2, ipady=2, side=tk.LEFT, expand=False)
OrderQtyEntry = ttk.Entry(FrameSim, textvariable=OrderQtyVar)
OrderQtyEntry.pack(side=tk.LEFT, expand=False)

# add a button to save new policy using the data in the Entry boxes
PolicyChangeButton = tk.Button(FrameSim, text='Change Policy', command=ChangePolicy).pack(padx=5, ipadx=2, ipady=2,
                                                                                          side=tk.RIGHT, expand=False)


# Step 6:
# Upload forecast, lead time, and review period from Temporary Excel file -----------
def UploadInputData():
    global df1
    InputFileName = fd.askopenfilename()
    try:
        ExcelFileInput = pd.ExcelFile(InputFileName)
    except:
        showinfo(title="File open", message="Error is found for Excel file opening.")
    else:
        df1 = pd.read_excel(ExcelFileInput, "Forecast and LT")

        SKUCode = df1["SKU Code"]
        SKUName = df1["SKU Name"]
        Forecast01 = df1["Month 1"]
        Forecast02 = df1["Month 2"]
        Forecast03 = df1["Month 3"]
        Forecast04 = df1["Month 4"]
        Forecast05 = df1["Month 5"]
        Forecast06 = df1["Month 6"]
        Forecast07 = df1["Month 7"]
        Forecast08 = df1["Month 8"]
        Forecast09 = df1["Month 9"]
        Forecast10 = df1["Month 10"]
        Forecast11 = df1["Month 11"]
        Forecast12 = df1["Month 12"]
        MinLT = df1["Leadtime Min"]
        ModeLT = df1["Leadtime Mode"]
        MaxLT = df1["Leadtime Max"]
        ReviewPeriod = df1["Review Period"]

        # display Forecast
        for i in range(TotalSKU):
            ViewForecast.item(i + 1, text=SKUCode[i],
                              values=(i + 1, SKUCode[i], SKUName[i], Forecast01[i], Forecast02[i], \
                                      Forecast03[i], Forecast04[i], Forecast05[i], Forecast06[i], \
                                      Forecast07[i], Forecast08[i], Forecast09[i], Forecast10[i], Forecast11[i],
                                      Forecast12[i]))
        # display Lead time and review period
        for i in range(TotalSKU):
            ViewLT.item(i + 1, text=SKUCode[i],
                        values=(i + 1, SKUCode[i], SKUName[i], MinLT[i], ModeLT[i], MaxLT[i], ReviewPeriod[i]))

        # show message
        showinfo(title='Data uploading', message='Forecast and Lead time data have been uploaded successfully!')


# Step 7:
# upload preliminary policy and KPI targets from data 06 file ---------------------------------
def UploadPolicy():
    global dfPolicy
    PolicyFileName = fd.askopenfilename()
    try:
        PolityInput = pd.ExcelFile(PolicyFileName)
    except Exception as e:
        showinfo(title="File open", message=e)
    else:
        dfPolicy = pd.read_excel(PolityInput, "Target Inventory Position")

        SKUCode = dfPolicy["SKU Code"]
        SKUName = dfPolicy["SKU Name"]
        TargetTurns = dfPolicy["Target Yearly Turns"]
        TargetFillRate = dfPolicy["Target Fill Rate"]
        TargetIP = dfPolicy["Target Inventory Position"]
        ROP = dfPolicy["Initial ROP"]
        OrderQty = dfPolicy["Initial Order Qty"]

        # display performance targets
        for i in range(TotalSKU):
            ViewTarget.item(i + 1, text=SKUCode[i], values=(i + 1, SKUCode[i], SKUName[i], TargetTurns[i], \
                                                            TargetFillRate[i], TargetIP[i]))
        # display policy
        for i in range(TotalSKU):
            ViewSim.item(i + 1, text=SKUCode[i], values=(i + 1, SKUCode[i], SKUName[i], ROP[i], \
                                                         OrderQty[i], "", "", "", "", ""))

        # show message
        showinfo(title='Data uploading', message='Policy data have been uploaded successfully!')


# Step 8:
# 8.1 generate weekly forecast randomly ---------------------------------------

def GenerateWeeklyForecast(MonthlyForecast):
    # input: forecast for 12 months
    # return: weekly forecast for 12x4 = 48  weeks
    MinValue = 1
    MaxValue = 10

    WeeklyForecast = pd.Series(47)
    for i in range(12):
        PercentWeek01 = random.uniform(MinValue, MaxValue)
        PercentWeek02 = random.uniform(MinValue, MaxValue)
        PercentWeek03 = random.uniform(MinValue, MaxValue)
        PercentWeek04 = random.uniform(MinValue, MaxValue)
        Total = PercentWeek01 + PercentWeek02 + PercentWeek03 + PercentWeek04

        Value1 = round(MonthlyForecast.iloc[i] * (PercentWeek01 / Total))
        Value2 = round(MonthlyForecast.iloc[i] * (PercentWeek02 / Total))
        Value3 = round(MonthlyForecast.iloc[i] * (PercentWeek03 / Total))
        Value4 = round(MonthlyForecast.iloc[i]) - Value1 - Value2 - Value3

        WeeklyForecast[i * 4] = Value1
        WeeklyForecast[i * 4 + 1] = Value2
        WeeklyForecast[i * 4 + 2] = Value3
        WeeklyForecast[i * 4 + 3] = Value4

    return WeeklyForecast, round(WeeklyForecast.mean())


# 8.2 generate supply lead time randomly --------------------------------------
def GenerateSupplyLT(LTMin, LTMode, LTMax):
    NewLT = random.triangular(LTMin, LTMax, LTMode)

    return round(NewLT)


# Step 9:

# Single simulation run -------------------------------------------------------
def SingleRun(SKUIndex):
    # input:
    # 1) SKUIndex:
    #       used to retrieve forecast, supply lead time, review period,from df1
    #       used to retrieve ROP and OrderQty from ViewSim
    # Returns:
    #   1) OK or Error
    #   2) Yearly Demand
    #   3) Yearly Lost Sales
    #   4) Weekly IP
    #   5) Weekly On Hand
    #   6) Weekly Open Order
    #   7) Weekly Fill RateS
    #   8) Yearly Turns

    global dfPolicy
    global df1

    try:
        # Step 0: -----------------------------------------
        # retrieve Policy from ViewSim
        Record = ViewSim.item(SKUIndex + 1)
        Record_col = Record['values']

        ROP = Record_col[3]
        OrderQty = Record_col[4]
        print(f"Running simulation for SKU: {SKUIndex}, ROP: {ROP}, Order Qty: {OrderQty}")

        # retrieve forecast, supply lead times, Review period from df1
        MonthlyForecast = df1.iloc[SKUIndex, 3:15]
        LTMin = df1.iat[SKUIndex, 15]
        LTMode = df1.iat[SKUIndex, 16]
        LTMax = df1.iat[SKUIndex, 17]
        MyReviewPeriod = df1.iat[SKUIndex, 18]

        # Step 1: Create a DataFrame "dfSingleRun" for single simulation run --------
        NumRow = 8
        NumCol = 80
        # from Col 1, first 10 weeks as warmup period,
        #        next 48 weeks used for performance study, next 20 weeks as tail
        KeyOfRow = ["Weekly Demand", "Weekly Lost Sales", "Arrival (beginning)", "On Hand (beginning)",
                    "Open Order (end)", "Inventory Position (end)", "Purchase Qty", "Supply LT"]

        ColData = [0]
        for i in range(NumRow - 1):
            ColData.append(0)

        DataDict = {x: ColData for x in range(NumCol)}

        # convert the matrix into DataFrame
        dfSingleRun = pd.DataFrame(DataDict, index=KeyOfRow)

        # row 0: Weekly Demand
        # row 1: weekly lost Sales
        # row 2: Arrival
        # row 3: On Hand
        # row 4: Open order
        # row 5: Inventory Position
        # row 6: Purchase Qty
        # row 7: Supply LT
        # Col 0: initial on hand
        # Col 1: warm up
        # from Col 10: 48 weeks period for performance study
        # then remaining 20 weeks as tail, just to keep Open Order if any
        # Step 2: set initial values: forecast and on hand at beginning --------
        WarmUp = 10
        StudyPeriod = 48

        # assign on-hand initial values
        dfSingleRun.iat[3, 1] = round(ROP * 1.3)

        # assign weekly demand initial values for warmup [1 to 10] and normal period [11 to 58]
        WeeklyForecast, WeeklyForecastMean = GenerateWeeklyForecast(MonthlyForecast)

        for i in range(1, WarmUp + 1, 1):
            dfSingleRun.iat[0, i] = WeeklyForecastMean

        for i in range(11, WarmUp + StudyPeriod + 1, 1):
            dfSingleRun.iat[0, i] = WeeklyForecast.iat[i - 11]
        # Step 3: at the end of each week
        for WeekIndex in range(1, WarmUp + StudyPeriod + 1, 1):
            Demand = dfSingleRun.iat[0, WeekIndex]
            CurrTotalAvailable = dfSingleRun.iat[2, WeekIndex] + dfSingleRun.iat[3, WeekIndex]
            # update "Lost Sales" for current week, and "On-hand" for next week
            if (CurrTotalAvailable - Demand) >= 0:
                dfSingleRun.iat[3, WeekIndex + 1] = (CurrTotalAvailable - Demand)
                dfSingleRun.iat[1, WeekIndex] = 0
            else:
                dfSingleRun.iat[3, WeekIndex + 1] = 0
                dfSingleRun.iat[1, WeekIndex] = (Demand - CurrTotalAvailable)

            # update IP = on-hand at beginning next week + current open orders
            CurrOpenOrder = 0
            for r in range(WeekIndex, 79, 1):
                CurrOpenOrder = CurrOpenOrder + dfSingleRun.iat[2, r + 1]
            dfSingleRun.iat[4, WeekIndex] = CurrOpenOrder

            CurrIP = dfSingleRun.iat[3, WeekIndex + 1] + CurrOpenOrder
            dfSingleRun.iat[5, WeekIndex] = CurrIP
        # Step 4: Purchase decision
            if WeekIndex % MyReviewPeriod == 0:
                if CurrIP <= ROP:  # raise Purchase Request
                    # get a random supply LT
                    MyLT = GenerateSupplyLT(LTMin, LTMode, LTMax)
                    # Place an order --> Open Order
                    dfSingleRun.iat[2, WeekIndex + MyLT + 1] = OrderQty
                    dfSingleRun.iat[6, WeekIndex] = OrderQty
                    dfSingleRun.iat[7, WeekIndex] = MyLT
        # 9.5 Calculate Performance statistics ----------------------------
        YearlyDemand = 0
        YearlyLostSales = 0
        WeeklyIP = 0
        WeeklyOnHand = 0
        WeeklyOpenOrder = 0

        for WeekIndex in range(WarmUp + 1, WarmUp + StudyPeriod + 1, 1):
            YearlyDemand = YearlyDemand + dfSingleRun.iat[0, WeekIndex]
            YearlyLostSales = YearlyLostSales + dfSingleRun.iat[1, WeekIndex]

            WeeklyOnHand = WeeklyOnHand + dfSingleRun.iat[3, WeekIndex]
            WeeklyOpenOrder = WeeklyOpenOrder + dfSingleRun.iat[4, WeekIndex]
            WeeklyIP = WeeklyIP + dfSingleRun.iat[5, WeekIndex]

        # Calculate weekly values
        WeeklyOnHand = round(WeeklyOnHand / StudyPeriod)
        WeeklyOpenOrder = round(WeeklyOpenOrder / StudyPeriod)
        WeeklyIP = round(WeeklyIP / StudyPeriod)
        YearlyFillRate = (YearlyDemand - YearlyLostSales) / YearlyDemand
        YearlyTurns = YearlyDemand / WeeklyIP

        # export dfSingleRun to Excel file for validation (for selected SKUs)
        # print ("--- save single run result to file --------")
        # if (SKUIndex in [0, 1, 10, 18,19]):
        #     with pd.ExcelWriter("SKU" + str(SKUIndex + 1) + ".xlsx") as writer:
        #         dfSingleRun.to_excel(writer, sheet_name="SKU" + str(SKUIndex + 1))
        print(
            f"Simulation output for SKU1029 - Demand: {YearlyDemand}, Lost Sales: {YearlyLostSales}, IP: {WeeklyIP},"
            f" On Hand: {WeeklyOnHand}, Open Order: {WeeklyOpenOrder}, Fill Rate: {YearlyFillRate}, Turns: {YearlyTurns}")
        # 9.6 return results --------------------------------------------
        # 1) OK or Error
        # 2) Yearly Demand
        # 3) Yearly Lost Sales
        # 4) Weekly IP
        # 5) Weekly On Hand
        # 6) Weekly Open Order
        # 7) Yearly Fill Rate
        # 8) Yearly Turns
        return "Ok", YearlyDemand, YearlyLostSales, WeeklyIP, WeeklyOnHand, \
            WeeklyOpenOrder, YearlyFillRate, YearlyTurns

    except IndexError as e:
        print(f"Indexing error occurred: {e}")
    except KeyError as e:
        print(f"Key error occurred: {e}")
    except ValueError as e:
        print(f"Value error occurred: {e}")
    except ZeroDivisionError as e:
        print(f"Division by zero error occurred: {e}")
    except Exception as e:
        # Catch any other unhandled exception
        print(f"An unexpected error occurred: {e}")
    finally:
        print("Simulation completed for SKUIndex:", SKUIndex)

    return "Error", 0, 0, 0, 0, 0, 0, 0

# Run simulation function for all SKUs-----------------------------------------
def RunSimulation():
    # generate performance simulation for all SKUs
    # For a single SKU:
    #   call SingleRun () get performance for each run
    #   calculate SKU statistical performance based the results of multiple runs
    #   Disply simulated performance on GUI

    TotalRun = 50
    dfPerformance = pd.DataFrame()

    # Create a DataFrame to store simulation results for multiple runs (Run No = TotalRun)
    ColData = [0]
    for i in range(TotalRun - 1):
        ColData.append(0)
    ColSeries = pd.Series(ColData)
    PerformanceDict = {"Yearly Demand": ColSeries, "Yearly Lost Sales": ColSeries, \
                       "Weekly IP": ColSeries, "Weekly On Hand": ColSeries, \
                       "Weekly Open Order": ColSeries, \
                       "Yearly Fill Rate": ColSeries, "Yearly Turns": ColSeries}
    dfPerformance = pd.DataFrame(PerformanceDict)

    # Step 0: clear up performance on ViewSim ------------------------------------------------
    for i in range(TotalSKU):
        Record = ViewSim.item(i + 1)
        Record_col = Record['values']

        CurrIndex = Record_col[0]
        SKUCode = Record_col[1]
        SKUName = Record_col[2]
        ROP = Record_col[3]
        OrderQty = Record_col[4]

        ViewSim.item(CurrIndex, text=SKUCode, values=(CurrIndex, SKUCode, SKUName, \
                                                      ROP, OrderQty, "", "", "", "", ""))

    # Simulate performance for all SKUs
    for SKUIndex in range(TotalSKU):

        # for each SKU
        # Step 1: Clear up the dePerformance DataFrame for new SKU ------------
        dfPerformance['Yearly Demand'] = 0
        dfPerformance['Yearly Lost Sales'] = 0
        dfPerformance['Weekly IP'] = 0
        dfPerformance['Weekly On Hand'] = 0
        dfPerformance['Weekly Open Order'] = 0
        dfPerformance['Yearly Fill Rate'] = 0
        dfPerformance['Yearly Turns'] = 0

        # for each individual SKU, collect the simulation result from multiple runs
        ResultIndex = 0
        for RunIndex in range(TotalRun):

            # Step 2: call single run -----------------------------------------
            SingleRunResult, YearlyDemand, YearlyLostSales, WeeklyIP, WeeklyOnHand, WeeklyOpenOrder, YearlyFillRate, YearlyTurns = SingleRun(SKUIndex)
            # record single run result
            if SingleRunResult == "Ok":
                dfPerformance.iat[RunIndex, 0] = YearlyDemand
                dfPerformance.iat[RunIndex, 1] = YearlyLostSales
                dfPerformance.iat[RunIndex, 2] = WeeklyIP
                dfPerformance.iat[RunIndex, 3] = WeeklyOnHand
                dfPerformance.iat[RunIndex, 4] = WeeklyOpenOrder
                dfPerformance.iat[RunIndex, 5] = YearlyFillRate
                dfPerformance.iat[RunIndex, 6] = YearlyTurns
                ResultIndex = ResultIndex + 1
            else:
                # Error was found in Single Run,
                # Stop simulation for the particular SKU
                break

        if ResultIndex < TotalRun:  # skip the SKU
            continue

            # Export DfPerformance DataFrame to file for validation
        # dfPerformance.to_excel (SKUCode + " Performance.xlsx", sheet_name="dfPerformance")

        # Step 3: summarize performance statstics for multiple runs (=TotalRun)-----
        GUIYearlyTurns = round(dfPerformance['Yearly Turns'].mean(), 2)
        GUIYearlyFillRate = format(dfPerformance['Yearly Fill Rate'].mean(), ".00%")
        GUIWeeklyIP = dfPerformance['Weekly IP'].mean()
        GUIWeeklyOnHand = dfPerformance['Weekly On Hand'].mean()
        GUIWeeklyOpenOrder = dfPerformance['Weekly Open Order'].mean()

        # Step 4: Display the simulated performance on GUI (ViewSim)-----------
        Record = ViewSim.item(SKUIndex + 1)
        Record_col = Record['values']

        CurrIndex = Record_col[0]
        SKUCode = Record_col[1]
        SKUName = Record_col[2]
        ROP = Record_col[3]
        OrderQty = Record_col[4]

        print(" -- Complete simulation run for: " + SKUCode)
        print("ROP:" + str(ROP) + "     OrderQty:" + str(OrderQty))
        print(
            f"Updating ViewSim for SKU {SKUCode} with Turns: {GUIYearlyTurns}, Fill Rate: {GUIYearlyFillRate}, "
            f"Weekly IP: {GUIWeeklyIP}, On Hand: {GUIWeeklyOnHand}, Open Order: {GUIWeeklyOpenOrder}")

        ViewSim.item(CurrIndex, text=SKUCode, values=(CurrIndex, SKUCode, SKUName,
                                                      ROP, OrderQty, GUIYearlyTurns, GUIYearlyFillRate,
                                                      GUIWeeklyIP, GUIWeeklyOnHand, GUIWeeklyOpenOrder))
    print(" -------- Simulation END --------")


# Step 10:
# Export result to Excel file
def ExportResult():
    # export ViewSim to data file 07:
    # Excel file name: 'Project Data07 Performance prediction.xlsx'
    # worksheet name: 'Policy and Performance'

    showinfo(title='Data Exporting', message='Save to a local file: Project Data07 Performance prediction.xlsx')

    # read columns from ViewSim
    MyColumns = ViewSim["columns"]

    # define a new empty DataFrame
    dfPerformance = pd.DataFrame(columns=MyColumns)

    print("---- policy and predicted performance ----------")
    for i in range(1, TotalSKU + 1, 1):
        MyRecord = ViewSim.item(i)
        MyValues = MyRecord['values']

        dfPerformance.loc[i] = MyValues

    print(dfPerformance)
    dfPerformance.to_excel("Project Data07 Performance prediction.xlsx", sheet_name="Policy and Performance",
                           index=False)


# menu list
MainMenu = tk.Menu(MainWindow)
MainWindow.config(menu=MainMenu)
FileMenu = tk.Menu(MainMenu)
MainMenu.add_cascade(label='File', menu=FileMenu)
FileMenu.add_command(label='Open forecast & LT (Temp File)', command=UploadInputData)
FileMenu.add_command(label='Open Preliminary Policy (Data File 06)', command=UploadPolicy)
FileMenu.add_command(label='Save Result (To Data File 07)', command=ExportResult)
FileMenu.add_separator()
FileMenu.add_command(label='Exit', command=MainWindow.destroy)

FunctionMenu = tk.Menu(MainMenu)
MainMenu.add_cascade(label='Function', menu=FunctionMenu)
FunctionMenu.add_command(label='Run Simulation', command=RunSimulation)

# ---------------------------
MainWindow.mainloop()
# ---------------------------
