# Step 0:
# Create the Main GUI Window --------------------------------------------------

import tkinter as tk
import random
import pandas as pd
import numpy as np

from tkinter import TclError, ttk, Text
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showerror, showwarning, showinfo


from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from pandas.tseries.offsets import DateOffset


# Main Window
MainWindow = tk.Tk()
MainWindow.title('Inventory Performance Simulation')
MainWindow.geometry("910x450+50+50")
MainWindow.resizable(True, True)

# Global variable with default value = 4 weeks
TotalSKU = 20
df1 = pd.DataFrame()
dfPolicy = pd.DataFrame()

ROPVar = tk.StringVar()
OrderQtyVar = tk.StringVar()
Record_values = pd.Series()

# Step 1:
# Create NoteBook  ------------------------------------------------------------

MyNotebook = ttk.Notebook(MainWindow)
MyNotebook.pack(pady=10, expand=True)

# Create frames
FrameForecast = ttk.Frame(MyNotebook, width=900, height=400)
FrameLT = ttk.Frame(MyNotebook, width=900, height=400)
FrameTarget = ttk.Frame(MyNotebook, width=900, height=400)
FrameSim = ttk.Frame(MyNotebook, width=700, height=400)

FrameForecast.pack(fill='both', expand=True)
FrameLT.pack(fill='both', expand=True)
FrameTarget.pack(fill='both', expand=True)
FrameSim.pack(fill=tk.Y, expand=True, anchor=tk.E)

# Add frames to notebook
MyNotebook.add(FrameForecast, text='Demand Forecast')
MyNotebook.add(FrameLT, text='Lead Time & Review Period')
MyNotebook.add(FrameTarget, text='Performance Target')
MyNotebook.add(FrameSim, text="Performance Evaluation")

# Step 2:
# Create treeview for forecast (input) ----------------------------------------

ColumnForecast = ('SerialNo', 'SKU Code', 'SKU Name', "Month 1", "Month 2", "Month 3", "Month 4",
                  "Month 5", "Month 6", "Month 7", "Month 8", "Month 9", "Month 10",
                  "Month 11", "Month 12")
ViewForecast = ttk.Treeview(FrameForecast, columns=ColumnForecast, height=15, show='headings')

for col in ColumnForecast:
    ViewForecast.column(col, width=80, anchor=tk.CENTER)
    ViewForecast.heading(col, text=col)

for i in range(TotalSKU):
    ViewForecast.insert('', tk.END, iid=i + 1, text="", values=(i + 1, "", "", "", "", "", "", "",
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

for col in ColumnLT:
    ViewLT.column(col, width=80, anchor=tk.CENTER)
    ViewLT.heading(col, text=col)

for i in range(TotalSKU):
    ViewLT.insert('', tk.END, iid=i + 1, text="", values=(i + 1, "", "", "", "", "", "", ""))

ViewLT.pack(ipadx=20, ipady=20, anchor=tk.CENTER, fill=tk.X, expand=True)

# Step 4:
# Create treeview for performance target ---------------------------------------

ColumnTarget = (
    'SerialNo', 'SKU Code', 'SKU Name', "Target Yearly Turns", "Target Fill Rate", "Target Inventory Position")

ViewTarget = ttk.Treeview(FrameTarget, columns=ColumnTarget, height=15, show='headings')

for col in ColumnTarget:
    ViewTarget.column(col, width=100, anchor=tk.CENTER)
    ViewTarget.heading(col, text=col)

for i in range(TotalSKU):
    ViewTarget.insert('', tk.END, iid=i + 1, text="", values=(i + 1, "", "", "", "", "", ""))

ViewTarget.pack(ipadx=20, ipady=20, anchor=tk.CENTER, fill=tk.BOTH, expand=True)

# Step 5:
# Create treeview for simulation result ---------------------------------------

ColumnSim = ('No', 'SKU Code', 'SKU Name', "ROP", "Order Qty", "Predicted Turns",
             "Predicted Fill Rate", "Predicted Inventory Position",
             "Predicted On Hand", "Predicted Open Order")
ViewSim = ttk.Treeview(FrameSim, columns=ColumnSim, height=15, show='headings')

for col in ColumnSim:
    ViewSim.column(col, width=100, anchor=tk.CENTER)
    ViewSim.heading(col, text=col)

for i in range(TotalSKU):
    ViewSim.insert('', tk.END, iid=i + 1, text="", values=(i + 1, "", "", "", "", "", "", "", "", "", ""))

ViewSim.pack(ipadx=20, ipady=20, anchor=tk.E, fill=tk.BOTH, expand=True)


# Modify policy for selected item
def ChangePolicy():
    global Record_values
    global ROPVar
    global OrderQtyVar

    try:
        CurrIndex = Record_values[0]
        SKUCode = Record_values[1]
        SKUName = Record_values[2]

        # Validate input before proceeding
        ROP = int(ROPVar.get())
        OrderQty = int(OrderQtyVar.get())

        # Update the new policy in the simulation view
        ViewSim.item(CurrIndex, text=SKUCode, values=(
            CurrIndex, SKUCode, SKUName, ROP, OrderQty, "", "", "", "", ""
        ))
        showinfo(title="policy", message="Policy has been updated!")

    except ValueError:
        showerror(title="Input Error", message="Please enter valid integers for ROP and Order Quantity.")
    except IndexError:
        showerror(title="Selection Error", message="No record is selected.")
    except Exception as e:
        showerror(title="Error", message=f"An unexpected error occurred: {str(e)}")


# Record selected item and show its policy to the entry boxes
def ResultItem_selected(event):
    global ROPVar
    global OrderQtyVar
    global Record_values

    try:
        Selected_Record_No = ViewSim.selection()
        if not Selected_Record_No:
            showinfo(title="Error", message="No item selected.")
            return

        Record = ViewSim.item(Selected_Record_No)
        Record_values = Record['values']
        ROP = Record_values[3]
        OrderQty = Record_values[4]

        ROPEntry.delete(0, 'end')
        OrderQtyEntry.delete(0, 'end')
        ROPEntry.insert(0, str(ROP))
        OrderQtyEntry.insert(0, str(OrderQty))

    except TclError as e:
        showinfo(title="Error", message=f"TclError: {str(e)}")
    except IndexError:
        showinfo(title="Error", message="Error accessing selected record values.")
    except Exception as e:
        showinfo(title="Error", message=f"Unexpected error: {str(e)}")


ViewSim.bind('<<TreeviewSelect>>', ResultItem_selected)

# Add two labels and two Entry boxes
ttk.Label(FrameSim, text="New ROP:").pack(padx=5, ipadx=2, ipady=2, side=tk.LEFT, expand=False)
ROPEntry = ttk.Entry(FrameSim, textvariable=ROPVar)
ROPEntry.pack(side=tk.LEFT, expand=False)

ttk.Label(FrameSim, text="New Order Qty:").pack(padx=10, ipadx=2, ipady=2, side=tk.LEFT, expand=False)
OrderQtyEntry = ttk.Entry(FrameSim, textvariable=OrderQtyVar)
OrderQtyEntry.pack(side=tk.LEFT, expand=False)

# Add a button to save new policy using the data in the Entry boxes
PolicyChangeButton = tk.Button(FrameSim, text='Change Policy', command=ChangePolicy)
PolicyChangeButton.pack(padx=5, ipadx=2, ipady=2, side=tk.RIGHT, expand=False)


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
        for i in range(TotalSKU):
            ViewForecast.item(i + 1, text=SKUCode[i], values=(i + 1, SKUCode[i], SKUName[i],
                                                              df1.iloc[i]["Month 1"], df1.iloc[i]["Month 2"],
                                                              df1.iloc[i]["Month 3"], df1.iloc[i]["Month 4"],
                                                              df1.iloc[i]["Month 5"], df1.iloc[i]["Month 6"],
                                                              df1.iloc[i]["Month 7"], df1.iloc[i]["Month 8"],
                                                              df1.iloc[i]["Month 9"], df1.iloc[i]["Month 10"],
                                                              df1.iloc[i]["Month 11"], df1.iloc[i]["Month 12"]))
            ViewLT.item(i + 1, text=SKUCode[i], values=(i + 1, SKUCode[i], SKUName[i],
                                                        df1.iloc[i]["Leadtime Min"], df1.iloc[i]["Leadtime Mode"],
                                                        df1.iloc[i]["Leadtime Max"], df1.iloc[i]["Review Period"]))
        showinfo(title='Data uploading', message='Forecast and Lead time data have been uploaded successfully!')


# Step 7:
# Upload preliminary policy and KPI targets from data 06 file ---------------------------------
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
        for i in range(TotalSKU):
            ViewTarget.item(i + 1, text=SKUCode[i], values=(i + 1, SKUCode[i], SKUName[i],
                                                            dfPolicy.iloc[i]["Target Yearly Turns"],
                                                            dfPolicy.iloc[i]["Target Fill Rate"],
                                                            dfPolicy.iloc[i]["Target Inventory Position"]))
            ViewSim.item(i + 1, text=SKUCode[i], values=(i + 1, SKUCode[i], SKUName[i],
                                                         dfPolicy.iloc[i]["Initial ROP"],
                                                         dfPolicy.iloc[i]["Initial Order Qty"], "", "", "", "", ""))
        showinfo(title='Data uploading',
                 message='Initial policy and performance targets have been uploaded successfully!')


# Step 8:
# Generate weekly forecast randomly ---------------------------------------
def GenerateWeeklyForecast(MonthlyForecast):
    MinValue = 1
    MaxValue = 10
    WeeklyForecast = pd.Series([0] * 48)

    for i in range(12):
        PercentWeek01 = random.uniform(MinValue, MaxValue)
        PercentWeek02 = random.uniform(MinValue, MaxValue)
        PercentWeek03 = random.uniform(MinValue, MaxValue)
        PercentWeek04 = random.uniform(MinValue, MaxValue)
        Total = PercentWeek01 + PercentWeek02 + PercentWeek03 + PercentWeek04

        Value1 = round(MonthlyForecast[i] * (PercentWeek01 / Total))
        Value2 = round(MonthlyForecast[i] * (PercentWeek02 / Total))
        Value3 = round(MonthlyForecast[i] * (PercentWeek03 / Total))
        Value4 = round(MonthlyForecast[i]) - Value1 - Value2 - Value3

        WeeklyForecast[i * 4] = Value1
        WeeklyForecast[i * 4 + 1] = Value2
        WeeklyForecast[i * 4 + 2] = Value3
        WeeklyForecast[i * 4 + 3] = Value4

    return WeeklyForecast, round(WeeklyForecast.mean())


# 8.2 generate supply lead time randomly --------------------------------------
def GenerateSupplyLT(LTMin, LTMode, LTMax):
    return round(random.triangular(LTMin, LTMax, LTMode))


# Single simulation run -------------------------------------------------------
# Single simulation run -------------------------------------------------------
def SingleRun(SKUIndex):
    global df1
    global ViewSim

    try:
        # Retrieve SKU data from forecast, lead time, review period, ROP, and Order Qty.
        SKUData = df1.iloc[SKUIndex]

        # Extract the monthly forecast for the SKU.
        monthly_forecast = [
            SKUData["Month 1"], SKUData["Month 2"], SKUData["Month 3"], SKUData["Month 4"],
            SKUData["Month 5"], SKUData["Month 6"], SKUData["Month 7"], SKUData["Month 8"],
            SKUData["Month 9"], SKUData["Month 10"], SKUData["Month 11"], SKUData["Month 12"]
        ]

        # Generate the weekly forecast from the monthly forecast.
        weekly_forecast, average_weekly_demand = GenerateWeeklyForecast(monthly_forecast)

        # Retrieve the lead time values.
        LTMin = SKUData["Lead time Min"]
        LTMode = SKUData["Lead time Mode"]
        LTMax = SKUData["Lead time Max"]
        average_lead_time = (LTMin + LTMode + LTMax) / 3

        # Generate a supply lead time.
        supply_lead_time = GenerateSupplyLT(LTMin, LTMode, LTMax)

        # Calculate Safety Stock using a simple approximation.
        service_level = 1.65  # Approximate for 95% service level.
        safety_stock = service_level * (average_weekly_demand * np.sqrt(average_lead_time))

        # Retrieve the ROP and OrderQty from the simulation view.
        Record = ViewSim.item(SKUIndex + 1)
        ROP = int(Record['values'][3])
        OrderQty = int(Record['values'][4])

        # Initialize variables to track inventory position, on-hand, open orders, and lost sales.
        inventory_position = ROP
        on_hand = ROP
        open_order = 0
        lost_sales = 0
        total_demand = sum(monthly_forecast)

        # Simulation over 48 weeks.
        weekly_ip = []
        weekly_on_hand = []
        weekly_open_order = []

        for week in range(48):
            demand = weekly_forecast[week]

            # Check if we need to place a new order (ROP).
            if inventory_position <= ROP:
                open_order += OrderQty
                # Assume the order arrives after the lead time.
                delivery_week = week + supply_lead_time
                if delivery_week < 48:
                    weekly_open_order.append((delivery_week, OrderQty))

            # Check for arriving orders.
            arriving_orders = [order for delivery_week, order in weekly_open_order if delivery_week == week]
            if arriving_orders:
                on_hand += sum(arriving_orders)
                weekly_open_order = [
                    (delivery_week, order) for delivery_week, order in weekly_open_order if delivery_week != week
                ]

            # Fulfill demand from on-hand inventory.
            if on_hand >= demand:
                on_hand -= demand
            else:
                # Calculate lost sales if demand exceeds on-hand inventory.
                lost_sales += demand - on_hand
                on_hand = 0

            # Update the inventory position.
            inventory_position = on_hand + sum(order for _, order in weekly_open_order)

            # Track the weekly metrics.
            weekly_ip.append(inventory_position)
            weekly_on_hand.append(on_hand)

        # Calculate the yearly fill rate.
        yearly_fill_rate = np.nan_to_num((total_demand - lost_sales) / total_demand) if total_demand > 0 else 1.0

        # Calculate the yearly turns (inventory turnover).
        average_inventory = np.nan_to_num(np.mean(weekly_ip)) if weekly_ip else 0
        yearly_turns = np.nan_to_num(total_demand / average_inventory) if average_inventory > 0 else 0

        # Calculate the average weekly values.
        avg_weekly_ip = round(np.nan_to_num(np.mean(weekly_ip)), 2) if weekly_ip else 0
        avg_weekly_on_hand = round(np.nan_to_num(np.mean(weekly_on_hand)), 2) if weekly_on_hand else 0
        avg_weekly_open_order = round(
            np.nan_to_num(np.mean([order for _, order in weekly_open_order])), 2
        ) if weekly_open_order else 0

        # Return the simulation results including Safety Stock and Inventory Turnover.
        return (
            "OK",  # Indicating successful run.
            total_demand,  # Yearly Demand.
            lost_sales,  # Yearly Lost Sales.
            avg_weekly_ip,  # Average Weekly Inventory Position.
            avg_weekly_on_hand,  # Average Weekly On-Hand Inventory.
            avg_weekly_open_order,  # Average Weekly Open Orders.
            yearly_fill_rate,  # Yearly Fill Rate.
            yearly_turns,  # Inventory Turnover.
            round(safety_stock, 2)  # Safety Stock value.
        )

    except Exception as e:
        # Handle any errors during the simulation and return an error message.
        print(f"Error in SingleRun for SKU Index {SKUIndex}: {str(e)}")
        return "Error", 0, 0, 0, 0, 0, 0.0, 0, 0


# Run simulation function for all SKUs-----------------------------------------
def RunSimulation():
    TotalRun = 50
    dfPerformance = pd.DataFrame()

    ColData = [0] * TotalRun
    PerformanceDict = {
        "Yearly Demand": ColData,
        "Yearly Lost Sales": ColData,
        "Weekly IP": ColData,
        "Weekly On Hand": ColData,
        "Weekly Open Order": ColData,
        "Yearly Fill Rate": ColData,
        "Yearly Turns": ColData,
        "Safety Stock": ColData
    }
    dfPerformance = pd.DataFrame(PerformanceDict)

    # Step 0: clear up performance on ViewSim
    for i in range(TotalSKU):
        Record = ViewSim.item(i + 1)
        Record_col = Record['values']
        CurrIndex = Record_col[0]
        SKUCode = Record_col[1]
        SKUName = Record_col[2]
        ROP = Record_col[3]
        OrderQty = Record_col[4]

        ViewSim.item(CurrIndex, text=SKUCode, values=(CurrIndex, SKUCode, SKUName,
                                                      ROP, OrderQty, "", "", "", "", ""))

    # Simulate performance for all SKUs
    for SKUIndex in range(TotalSKU):
        # Step 1: Clear up the dfPerformance DataFrame for new SKU
        dfPerformance.loc[:, :] = 0

        # Collect simulation results for each run
        ResultIndex = 0
        for RunIndex in range(TotalRun):
            SingleRunResult, YearlyDemand, YearlyLostSales, WeeklyIP, WeeklyOnHand, \
                WeeklyOpenOrder, YearlyFillRate, YearlyTurns, SafetyStock = SingleRun(SKUIndex)

            # Record single run result
            if SingleRunResult == "OK":
                dfPerformance.iloc[RunIndex] = [
                    YearlyDemand, YearlyLostSales, WeeklyIP, WeeklyOnHand,
                    WeeklyOpenOrder, YearlyFillRate, YearlyTurns, SafetyStock
                ]
                ResultIndex += 1
            else:
                # Error was found in Single Run, stop simulation for the particular SKU
                break

        # Skip SKU if not enough successful runs
        if ResultIndex < TotalRun:
            continue

        # Step 2: Summarize performance statistics for multiple runs
        GUIYearlyTurns = round(dfPerformance['Yearly Turns'].mean(), 2)
        GUIYearlyFillRate = format(dfPerformance['Yearly Fill Rate'].mean(), ".00%")
        GUIWeeklyIP = dfPerformance['Weekly IP'].mean()
        GUIWeeklyOnHand = dfPerformance['Weekly On Hand'].mean()
        GUIWeeklyOpenOrder = dfPerformance['Weekly Open Order'].mean()
        GUISafetyStock = round(dfPerformance['Safety Stock'].mean(), 2)

        # Step 3: Display the simulated performance on GUI (ViewSim)
        Record = ViewSim.item(SKUIndex + 1)
        Record_col = Record['values']
        CurrIndex = Record_col[0]
        SKUCode = Record_col[1]
        SKUName = Record_col[2]
        ROP = Record_col[3]
        OrderQty = Record_col[4]

        print(f" -- Complete simulation run for: {SKUCode}")
        print(
            f"ROP: {ROP}     OrderQty: {OrderQty}    SafetyStock: {GUISafetyStock}    Inventory Turnover: {GUIYearlyTurns}")

        # Ensure that the inventory turnover (GUIYearlyTurns) is shown in the "Predicted Turns" column
        ViewSim.item(CurrIndex, text=SKUCode, values=(
            CurrIndex, SKUCode, SKUName, ROP, OrderQty,
            GUIYearlyTurns, GUIYearlyFillRate,
            GUIWeeklyIP, GUIWeeklyOnHand, GUIWeeklyOpenOrder
        ))
    print(" -------- Simulation END --------")


# Step 10:
# Export result to Excel file
def ExportResult():
    showinfo(title='Data Exporting', message='Save to a local file: Project Data07 Performance prediction.xlsx')

    MyColumns = ViewSim["columns"]
    dfPerformance = pd.DataFrame(columns=MyColumns)

    print("---- policy and predicted performance ----------")
    for i in range(1, TotalSKU + 1):
        MyRecord = ViewSim.item(i)
        MyValues = MyRecord['values']
        dfPerformance.loc[i] = MyValues

    print(dfPerformance)
    dfPerformance.to_excel("Project Data07 Performance prediction.xlsx", sheet_name="Policy and Performance",
                           index=False)


# Menu list
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

MainWindow.mainloop()
