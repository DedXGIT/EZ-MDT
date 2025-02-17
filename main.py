import tkinter as tk
from tkinter import ttk
import pandas as pd
import pyperclip 

# Load data from Excel file (replace 'data.xlsx' with your file path)
try:
    df = pd.read_excel('SASPROSTER.xlsx')
    dropdown_options = df['OFFICERS'].tolist()  # Assuming 'OFFICERS' is the column name
    dropdown_car_options = df['CARS'].tolist()  # Assuming 'CARS' is the column name
except Exception as e:
    print(f"Error loading Excel file: {e}")
    dropdown_options = ["Option 1", "Option 2", "Option 3"]  # Fallback options

# Create the main window
root = tk.Tk()
root.title("EASY MDT")

# Frame for Basic Section
basic_frame = ttk.LabelFrame(root, text="Basic")
basic_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")

# Type dropdown
type_label = ttk.Label(basic_frame, text="Type:")
type_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
type_var = tk.StringVar()
type_dropdown = ttk.Combobox(basic_frame, textvariable=type_var, values=["BANK", "STORE", "JEWELLERY"])
type_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
type_dropdown.set("BANK")  # Default value

# Bank dropdown (visible only if Type is "bank")
bank_label = ttk.Label(basic_frame, text="Which Bank:")
bank_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
bank_var = tk.StringVar()
bank_dropdown = ttk.Combobox(basic_frame, textvariable=bank_var, values=[
    "LEGION BANK", "HAWICK BANK", "DEL PERRO BANK", "GRT OCEAN BANK", "HARMONY BANK", "PALETO BANK" , "PACIFIC BANK"
])
bank_dropdown.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
bank_dropdown.set("LEGION BANK")  # Default value

# Store dropdown (visible only if Type is "store")
store_label = ttk.Label(basic_frame, text="Which Store:")
store_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
store_var = tk.StringVar()
store_dropdown = ttk.Combobox(basic_frame, textvariable=store_var, values=[
    "DAVIS LTD GASOLINE", "STRAWBERRY 24/7", "MURIETTA HEIGHTS LIQUOR", "LITTLE SEOUL GASOLINE", "VESPUCCI CANALS LIQUOR" , "MORNINGWOOD ROB'S LIQUOR" , "MIRROR PARK GASOLINE" , "VINEWOOD 24/7" , "TATAVIAM MOUNTAINS 24/7" , "BANHAM CANYON 24/7" , "RICHMAN GLEN GASOLINE" , "CHUMASH 24/7" , "HARMONY 24/7" , "GRAND SENORA LIQUOR" , "SANDY SHORES 24/7" , "GRAPESEED GASOLINE" , "MOUNT CHILLIAD 24/7" , "PALETO BAY 24/7"
])
store_dropdown.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
store_dropdown.set("Vinewood Store")  # Default value

# Function to toggle dropdown visibility based on Type
def toggle_dropdowns(*args):
    selected_type = type_var.get()
    if selected_type == "BANK":
        bank_label.grid()
        bank_dropdown.grid()
        store_label.grid_remove()
        store_dropdown.grid_remove()
    elif selected_type == "STORE":
        store_label.grid()
        store_dropdown.grid()
        bank_label.grid_remove()
        bank_dropdown.grid_remove()
    else:
        bank_label.grid_remove()
        bank_dropdown.grid_remove()
        store_label.grid_remove()
        store_dropdown.grid_remove()

type_var.trace_add("write", toggle_dropdowns)
toggle_dropdowns()  # Initial call to set visibility

# Frame for Scene Details
scene_frame = ttk.LabelFrame(root, text="Scene Details")
scene_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

# REPORTING OFFICER dropdown
reporting_officer_label = ttk.Label(scene_frame, text="REPORTING OFFICER:")
reporting_officer_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
reporting_officer_var = tk.StringVar()
reporting_officer_dropdown = ttk.Combobox(scene_frame, textvariable=reporting_officer_var, values=dropdown_options)
reporting_officer_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
reporting_officer_dropdown.set(dropdown_options[0])  # Default value

# MDT Creator dropdown
mdt_creator_label = ttk.Label(scene_frame, text="MDT Creator:")
mdt_creator_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
mdt_creator_var = tk.StringVar()
mdt_creator_dropdown = ttk.Combobox(scene_frame, textvariable=mdt_creator_var, values=dropdown_options)
mdt_creator_dropdown.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
mdt_creator_dropdown.set(dropdown_options[0])  # Default value

# Negotiator dropdown
negotiator_label = ttk.Label(scene_frame, text="Negotiator:")
negotiator_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
negotiator_var = tk.StringVar()
negotiator_dropdown = ttk.Combobox(scene_frame, textvariable=negotiator_var, values=dropdown_options)
negotiator_dropdown.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
negotiator_dropdown.set(dropdown_options[0])  # Default value

# Scene Command dropdown
scene_command_label = ttk.Label(scene_frame, text="Scene Command:")
scene_command_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
scene_command_var = tk.StringVar()
scene_command_dropdown = ttk.Combobox(scene_frame, textvariable=scene_command_var, values=dropdown_options)
scene_command_dropdown.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
scene_command_dropdown.set(dropdown_options[0])  # Default value

# Stayed back for hostage dropdown
stayed_back_label = ttk.Label(scene_frame, text="Stayed back for hostage:")
stayed_back_label.grid(row=4, column=0, padx=10, pady=5, sticky="w")
stayed_back_var = tk.StringVar()
stayed_back_dropdown = ttk.Combobox(scene_frame, textvariable=stayed_back_var, values=dropdown_options)
stayed_back_dropdown.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
stayed_back_dropdown.set(dropdown_options[0])  # Default value

# Frame for Chase Sequence
chase_frame = ttk.LabelFrame(root, text="Chase Sequence")
chase_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

# Primary dropdown
primary_label = ttk.Label(chase_frame, text="Primary:")
primary_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
primary_var = tk.StringVar()
primary_dropdown = ttk.Combobox(chase_frame, textvariable=primary_var, values=dropdown_options)
primary_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
primary_dropdown.set(dropdown_options[0])  # Default 
primary_dropdown.config(width=35)
##Primary Car
primary_car_dropdown = ttk.Combobox(chase_frame, textvariable=primary_var, values=dropdown_car_options)
primary_car_dropdown.grid(row=0, column=2, padx=10, pady=5, sticky="ew")
primary_car_dropdown.set(dropdown_options[0])  # Default value
primary_car_dropdown.config(width=35)

# Secondary dropdown
secondary_label = ttk.Label(chase_frame, text="Secondary:")
secondary_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
secondary_var = tk.StringVar()
secondary_dropdown = ttk.Combobox(chase_frame, textvariable=secondary_var, values=dropdown_options)
secondary_dropdown.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
secondary_dropdown.set(dropdown_options[0])  # Default value

# Tertiary dropdown
tertiary_label = ttk.Label(chase_frame, text="Tertiary:")
tertiary_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
tertiary_var = tk.StringVar()
tertiary_dropdown = ttk.Combobox(chase_frame, textvariable=tertiary_var, values=dropdown_options)
tertiary_dropdown.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
tertiary_dropdown.set(dropdown_options[0])  # Default value

# Parallel dropdown
parallel_label = ttk.Label(chase_frame, text="Parallel:")
parallel_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
parallel_var = tk.StringVar()
parallel_dropdown = ttk.Combobox(chase_frame, textvariable=parallel_var, values=dropdown_options)
parallel_dropdown.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
parallel_dropdown.set(dropdown_options[0])  # Default value

# 5th Unit dropdown
fifth_unit_label = ttk.Label(chase_frame, text="5th Unit:")
fifth_unit_label.grid(row=4, column=0, padx=10, pady=5, sticky="w")
fifth_unit_var = tk.StringVar()
fifth_unit_dropdown = ttk.Combobox(chase_frame, textvariable=fifth_unit_var, values=dropdown_options)
fifth_unit_dropdown.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
fifth_unit_dropdown.set(dropdown_options[0])  # Default value

# AirOne dropdown
airone_label = ttk.Label(chase_frame, text="AirOne:")
airone_label.grid(row=5, column=0, padx=10, pady=5, sticky="w")
airone_var = tk.StringVar()
airone_dropdown = ttk.Combobox(chase_frame, textvariable=airone_var, values=dropdown_options)
airone_dropdown.grid(row=5, column=1, padx=10, pady=5, sticky="ew")
airone_dropdown.set(dropdown_options[0])  # Default value

# Frame for Robbers Details
robbers_frame = ttk.LabelFrame(root, text="Robbers Details")
robbers_frame.grid(row=3, column=0, padx=10, pady=5, sticky="ew")

# Robbers Inside
robbers_inside_label = ttk.Label(robbers_frame, text="Robbers Inside:")
robbers_inside_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
robbers_inside_var = tk.IntVar()
robbers_inside_entry = ttk.Entry(robbers_frame, textvariable=robbers_inside_var)
robbers_inside_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

# Robbers Outside
robbers_outside_label = ttk.Label(robbers_frame, text="Robbers Outside:")
robbers_outside_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
robbers_outside_var = tk.IntVar()
robbers_outside_entry = ttk.Entry(robbers_frame, textvariable=robbers_outside_var)
robbers_outside_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

# Hostage
hostage_label = ttk.Label(robbers_frame, text="Hostage:")
hostage_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
hostage_var = tk.IntVar()
hostage_entry = ttk.Entry(robbers_frame, textvariable=hostage_var)
hostage_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

# Frame for Vehicle Details
vehicle_frame = ttk.LabelFrame(root, text="Vehicle Details")
vehicle_frame.grid(row=4, column=0, padx=10, pady=5, sticky="ew")

# Model
model_label = ttk.Label(vehicle_frame, text="Model:")
model_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
model_var = tk.StringVar()
model_entry = ttk.Entry(vehicle_frame, textvariable=model_var)
model_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

# Colour
colour_label = ttk.Label(vehicle_frame, text="Colour:")
colour_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
colour_var = tk.StringVar()
colour_entry = ttk.Entry(vehicle_frame, textvariable=colour_var)
colour_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

# Plate
plate_label = ttk.Label(vehicle_frame, text="Plate:")
plate_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
plate_var = tk.StringVar()
plate_entry = ttk.Entry(vehicle_frame, textvariable=plate_var)
plate_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

# Registered on
registered_label = ttk.Label(vehicle_frame, text="Registered on:")
registered_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
registered_var = tk.StringVar()
registered_entry = ttk.Entry(vehicle_frame, textvariable=registered_var)
registered_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")


# Function to generate the report
def generate_report():
    # Get all input values
    type_value = type_var.get()
    bank_value = bank_var.get() if type_value == "BANK" else store_var.get() if type_value == "STORE" else "JEWELLERY STORE"
    reporting_officer_value = reporting_officer_var.get()
    mdt_creator_value = mdt_creator_var.get()
    scene_command_value = scene_command_var.get()
    negotiator_value = negotiator_var.get()
    stayed_back_value = stayed_back_var.get()
    primary_value = primary_var.get()
    secondary_value = secondary_var.get()
    tertiary_value = tertiary_var.get()
    parallel_value = parallel_var.get()
    fifth_unit_value = fifth_unit_var.get()
    airone_value = airone_var.get()
    robbers_inside_value = robbers_inside_var.get()
    robbers_outside_value = robbers_outside_var.get()
    hostage_value = hostage_var.get()
    model_value = model_var.get()
    colour_value = colour_var.get()
    plate_value = plate_var.get()
    registered_value = registered_var.get()

    # Generate the report
    report = f"""
10-90 | {type_value} - {bank_value}
______________________________________________________________________________________________________________________________

REPORTING OFFICER : {reporting_officer_value}  

SCENE ASSIGNMENT :
MDT Creator: {mdt_creator_value}
Scene Command: {scene_command_value}
Negotiator: {negotiator_value}
Stayed Back For Hostage: {stayed_back_value}

INVOLVED IN PURSUIT : 
Filled By Scene command assigned officer
Primary: {primary_value}
Secondary: {secondary_value}
Tertiary: {tertiary_value}
Parallel: {parallel_value}
5th Unit: {fifth_unit_value}
Air1: {airone_value}
______________________________________________________________________________________________________________________________
 DETAILS & DEMANDS :
While patrolling, we received a report of an alarm going off at the {bank_value}. 
{mdt_creator_value} was assigned to create an incident report.

After setting up the perimeters around the area, we began with the negotiations. 
By interacting with the robbers we learned few below mentioned things:

Robbers Inside: {robbers_inside_value}
Robbers Outside: {robbers_outside_value}
Hostage: {hostage_value}

Vehicle Details
Model: {model_value}
Colour: {colour_value}
Plate: {plate_value}
Registered on: {registered_value}

Robbers were unidentified and in exchange of the hostage, their demand was Free Passage and No Spikes .
Once everyone was ready, scene command prepared a lineup for the pursuit. 
______________________________________________________________________________________________________________________________ 
CHASE :
Once everyone was ready, the chase started and they attempted to evade from police recklessly. The officers followed the suspects vehicle according to the sequence. The robbers kept roaming in the city and were damaging the public property.
    """

    # Display the report in a new window
    report_window = tk.Toplevel(root)
    report_window.title("Generated Report")
    report_text = tk.Text(report_window, wrap=tk.WORD, width=100, height=40)
    report_text.insert(tk.END, report)
    report_text.config(state=tk.DISABLED)  # Make the text read-only
    report_text.pack(padx=10, pady=10)

    def copy_to_clipboard():
        root.clipboard_clear()
        root.clipboard_append(report)
        root.update()

    copy_button = ttk.Button(report_window, text="Copy to Clipboard", command=copy_to_clipboard)
    copy_button.pack(pady=5)

# Submit button
submit_button = ttk.Button(root, text="Submit", command=generate_report)
submit_button.grid(row=5, column=0, columnspan=2, pady=10)


# Run the application
root.mainloop()