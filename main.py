# Import libraries and specific functions
import pandas as pd
import tkinter as tk
from tkinter import CENTER, NO, RIGHT, Y, Scrollbar, ttk
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure

# Import required variables from the financial statement
needed_variables = ["Zysk/(strata) netto", "AKTYWA OBROTOWE", "AKTYWA RAZEM", "Przychody ze sprzedaży",
                    "KAPITAŁ WŁASNY", "AKTYWA OBROTOWE", "ZOBOWIĄZANIA KRÓTKOTERMINOWE", "Należności handlowe",
                    "Koszty sprzedanych produktów, usług, towarów i materiałów",
                    "Zapasy", "ZOBOWIĄZANIA DŁUGOTERMINOWE"]

# Create dictionaries to store values for specific years
values2020 = {}
values2019 = {}
values2018 = {}
values2017 = {}

# Read Excel file
file = pd.read_excel("Szablon-projektu-na-Pythona.xlsx")

# Assign specific columns from Excel file to variables
Indicators = file[file.columns[0]]
year2020 = file[file.columns[1]]
year2019 = file[file.columns[2]]
year2018 = file[file.columns[3]]
year2017 = file[file.columns[4]]

# Create a loop to assign values to needed variables
i = 0
for row in Indicators:
    for header in needed_variables:
        if header in str(row):
            values2020[header] = year2020[i]
            values2019[header] = year2019[i]
            values2018[header] = year2018[i]
            values2017[header] = year2017[i]
    i += 1

# Write a function to calculate specific indicators for given years
def calculate(values_from_current_year, values_from_previous_year):
    indicators = {}
    indicators["ROA"] = round(values_from_current_year["Zysk/(strata) netto"] / (
            (values_from_current_year["AKTYWA RAZEM"] + values_from_previous_year["AKTYWA RAZEM"]) / 2), 2)
    indicators["ROE"] = round(values_from_current_year["Zysk/(strata) netto"] / (
            (values_from_current_year["KAPITAŁ WŁASNY"] + values_from_previous_year["KAPITAŁ WŁASNY"]) / 2), 2)
    indicators["current_liquidity_ratio"] = round(values_from_current_year["AKTYWA OBROTOWE"] / (
        (values_from_current_year["ZOBOWIĄZANIA KRÓTKOTERMINOWE"])), 2)
    indicators["quick_liquidity_ratio"] = round(
        (values_from_current_year["AKTYWA OBROTOWE"] - values_from_current_year["Zapasy"]) / (
            values_from_current_year["ZOBOWIĄZANIA KRÓTKOTERMINOWE"]), 2)
    indicators["immediate_liquidity_ratio"] = round(values_from_current_year["Należności handlowe"] / (
        values_from_current_year["ZOBOWIĄZANIA KRÓTKOTERMINOWE"]), 2)
    indicators["total_debt_ratio"] = round((
            (values_from_current_year["ZOBOWIĄZANIA DŁUGOTERMINOWE"] + values_from_current_year[
                "ZOBOWIĄZANIA KRÓTKOTERMINOWE"]) / values_from_current_year["AKTYWA RAZEM"]), 2)
    indicators["equity_to_assets_ratio"] = round(
        (values_from_current_year["KAPITAŁ WŁASNY"] / values_from_current_year["AKTYWA RAZEM"]), 2)
    indicators["ROS"] = round((values_from_current_year["Zysk/(strata) netto"] /
                              values_from_current_year["Przychody ze sprzedaży"]), 2)
    indicators["ASSETS/EQUITY"] = round((values_from_current_year["AKTYWA RAZEM"] /
                              values_from_current_year["KAPITAŁ WŁASNY"]), 2)
    indicators["TOTAL_SALES"] = round((values_from_current_year["Przychody ze sprzedaży"] /
                              values_from_current_year["AKTYWA RAZEM"]), 2)
    return indicators

# Calculate indicators for specific years
indicators_2020 = calculate(values2020, values2019)
indicators_2019 = calculate(values2019, values2018)
indicators_2018 = calculate(values2018, values2017)

# Create window
root = tk.Tk()
root.title('Financial Statement')
root.configure(background="#596")

# Table
style = ttk.Style()
# Header text style configuration
style.configure("Treeview.Heading", font=(None, 14))
# Text style configuration
style.configure("Treeview", font=(None, 12), rowheight=19)

# Columns
tree = ttk.Treeview(root, show='headings')

tree['columns'] = ('Indicator', '2020', '2019', '2018')
tree.column("Indicator", anchor=CENTER, width=330)
tree.column("2020", anchor=CENTER, width=70)
tree.column("2019", anchor=CENTER, width=70)
tree.column("2018", anchor=CENTER, width=70)

# Define headers
tree.heading('Indicator', text='Indicator')
tree.heading('2020', text='2020')
tree.heading('2019', text='2019')
tree.heading('2018', text='2018')

# Generate indicator values for specific years
indicators = []
for key in indicators_2020:
    indicators.append((f'{key}', f'{indicators_2020[key]}', f'{indicators_2019[key]}', f'{indicators_2018[key]}',))
    print(indicators)
# Add indicator values from specific years to the table
for ind in indicators:
    tree.insert('', tk.END, values=ind)
# Set table position in the window
tree.grid(row=0, column=1, sticky='nsew')

# Graph 1
t_1 = [values2017["Zysk/(strata) netto"], values2018["Zysk/(strata) netto"],
       values2019["Zysk/(strata) netto"], values2020["Zysk/(strata) netto"]]

fig = Figure(figsize=(5, 2), dpi=85)
fig.add_subplot(111).plot(["2017", "2018", "2019", "2020"], t_1)
fig.suptitle('Net Profit/(Loss)', fontsize=12)

for x, y in zip(["2017", "2018", "2019", "2020"], t_1):
    label = str(float("{:.2f}".format(y)) / 1000) + " thous."

    fig.gca().annotate(label,  # This is the text
                       (x, y),  # These are the coordinates for label placement
                       textcoords="offset points",  # How to position the text
                       xytext=(0, 9),  # Distance from text to (x,y) points
                       ha='center')

canvas = FigureCanvasTkAgg(fig, master=root)
canvas.draw()
canvas.get_tk_widget().grid(row=1, column=0)
current_values = fig.gca().get_yticks()
fig.gca().set_yticklabels(['{:.0f}'.format(x) for x in current_values])

# Graph 2
t_3 = [values2017["Koszty sprzedanych produktów, usług, towarów i materiałów"],
       values2018["Koszty sprzedanych produktów, usług, towarów i materiałów"],
       values2019["Koszty sprzedanych produktów, usług, towarów i materiałów"],
       values2020["Koszty sprzedanych produktów, usług, towarów i materiałów"]]

fig2 = Figure(figsize=(6, 2), dpi=85)
fig2.add_subplot(111).plot(["2017", "2018", "2019", "2020"], t_3)
fig2.suptitle('Cost of Goods, Services, Merchandise and Materials Sold', fontsize=12)

for x, y in zip(["2017", "2018", "2019", "2020"], t_3):
    label = str(float("{:.2f}".format(y)) / 1000) + " thous."

    fig2.gca().annotate(label,  # This is the text
                        (x, y),  # These are the coordinates for label placement
                        textcoords="offset points",  # How to position the text
                        xytext=(0, -7),  # Distance from text to (x,y) points
                        ha='center')

canvas = FigureCanvasTkAgg(fig2, master=root)
canvas.draw()
canvas.get_tk_widget().grid(row=1, column=1)
current_values = fig2.gca().get_yticks()
fig2.gca().set_yticklabels(['{:.0f}'.format(x) for x in current_values])

# Graph 3
t_2 = [values2017["Przychody ze sprzedaży"], values2018["Przychody ze sprzedaży"],
       values2019["Przychody ze sprzedaży"], values2020["Przychody ze sprzedaży"]]

fig1 = Figure(figsize=(5, 2), dpi=85)
fig1.add_subplot(111).plot(["2017", "2018", "2019", "2020"], t_2)
fig1.suptitle('Sales Revenue', fontsize=12)

for x, y in zip(["2017", "2018", "2019", "2020"], t_2):
    label = str(float("{:.2f}".format(y)) / 1000) + " thous."

    fig1.gca().annotate(label,  # This is the text
                        (x, y),  # These are the coordinates for label placement
                        textcoords="offset points",  # How to position the text
                        xytext=(0, 5),  # Distance from text to (x,y) points
                        ha='center')

canvas = FigureCanvasTkAgg(fig1, master=root)
canvas.draw()
canvas.get_tk_widget().grid(row=1, column=2)
current_values = fig1.gca().get_yticks()
fig1.gca().set_yticklabels(['{:.0f}'.format(x) for x in current_values])

# Create Du Pont chart
# Year 2020
label2020 = tk.Label(root, text="2020", background="#00A8E8", foreground="#fff", font=(None, 16))
label2020.grid(row=5, column=2, padx=18, pady=12)

wskps2020_label = tk.Label(root, text=f'ASSETS/EQUITY = {indicators_2020["quick_liquidity_ratio"]}',
                           background="#00A8E8", foreground="#fff", font=(None, 16))
wskps2020_label.grid(row=6, column=2, padx=18, pady=12)

wskpb2020_label = tk.Label(root, text=f'TOTAL_SALES = {indicators_2020["current_liquidity_ratio"]}',
                           background="#00A8E8", foreground="#fff", font=(None, 16))
wskpb2020_label.grid(row=7, column=2, padx=18, pady=12)

roe2020_label = tk.Label(root, text=f'ROE = {indicators_2020["ROE"]}', background="#00A8E8", foreground="#fff",
                         font=(None, 16))
roe2020_label.grid(row=8, column=2, padx=18, pady=12)

roa2020_label = tk.Label(root, text=f'ROA = {indicators_2020["ROA"]}', background="#00A8E8", foreground="#fff",
                         font=(None, 16))
roa2020_label.grid(row=9, column=2, padx=18, pady=12)

ros2020_label = tk.Label(root, text=f'ROS = {indicators_2020["ROS"]}', background="#00A8E8", foreground="#fff",
                         font=(None, 16))
ros2020_label.grid(row=10, column=2, padx=18, pady=12)

# Year 2019
label2019 = tk.Label(root, text="2019", background="#95eb34", foreground="#fff", font=(None, 16))
label2019.grid(row=5, column=1, padx=18, pady=12)

wskps2019_label = tk.Label(root, text=f'ASSETS/EQUITY = {indicators_2019["quick_liquidity_ratio"]}',
                           background="#95eb34", foreground="#fff", font=(None, 16))
wskps2019_label.grid(row=6, column=1, padx=18, pady=12)

wskpb2019_label = tk.Label(root, text=f'TOTAL_SALES = {indicators_2019["current_liquidity_ratio"]}',
                           background="#95eb34", foreground="#fff", font=(None, 16))
wskpb2019_label.grid(row=7, column=1, padx=18, pady=12)

roe2019_label = tk.Label(root, text=f'ROE = {indicators_2019["ROE"]}', background="#95eb34", foreground="#fff",
                         font=(None, 16))
roe2019_label.grid(row=8, column=1, padx=18, pady=12)

roa2019_label = tk.Label(root, text=f'ROA = {indicators_2019["ROA"]}', background="#95eb34", foreground="#fff",
                         font=(None, 16))
roa2019_label.grid(row=9, column=1, padx=18, pady=12)

ros2019_label = tk.Label(root, text=f'ROS = {indicators_2019["ROS"]}', background="#95eb34", foreground="#fff",
                         font=(None, 16))
ros2019_label.grid(row=10, column=1, padx=18, pady=12)

# Year 2018
label2018 = tk.Label(root, text="2018", background="#eb4034", foreground="#fff", font=(None, 16))
label2018.grid(row=5, column=0, padx=18, pady=12)

wskps2018_label = tk.Label(root, text=f'ASSETS/EQUITY = {indicators_2018["quick_liquidity_ratio"]}',
                           background="#eb4034", foreground="#fff", font=(None, 16))
wskps2018_label.grid(row=6, column=0, padx=18, pady=12)

wskpb2018_label = tk.Label(root, text=f'TOTAL_SALES = {indicators_2018["current_liquidity_ratio"]}',
                           background="#eb4034", foreground="#fff", font=(None, 16))
wskpb2018_label.grid(row=7, column=0, padx=18, pady=12)

roe2018_label = tk.Label(root, text=f'ROE = {indicators_2018["ROE"]}', background="#eb4034", foreground="#fff",
                         font=(None, 16))
roe2018_label.grid(row=8, column=0, padx=18, pady=12)

roa2018_label = tk.Label(root, text=f'ROA = {indicators_2018["ROA"]}', background="#eb4034", foreground="#fff",
                         font=(None, 16))
roa2018_label.grid(row=9, column=0, padx=18, pady=12)

ros2018_label = tk.Label(root, text=f'ROS = {indicators_2018["ROS"]}', background="#eb4034", foreground="#fff",
                         font=(None, 16))
ros2018_label.grid(row=10, column=0, padx=18, pady=12)

# Start the application
root.mainloop()