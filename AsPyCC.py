"""Created on the 3/9/2025
    Author: Engr. Fernando Zea, PhD student @ The University of Kansas (KU)
    Contact: fzea@ku.edu, fzea@espol.edu.ec, fernandozea98@gmail.com
    From the publication on Computers and Chemical Engieering: Pending
    DOI: Pending
    How to cite: Pending
"""

# Importing required libraries
import os
import win32com.client as win32
import numpy as np
import pandas as pd
import time

# Defining input data (This is defined by the user)
flue_gas_data = pd.read_csv(r'') 
df_flue_gas = flue_gas_data.loc[flue_gas_data['Industry'] == ''] # See Data_generation.ipynb for more information regarding database structure

# Initializing Aspen Plus simulation file
Aspen_file_path = r'' # A .bkp file is recommended 

# Access the Aspen Plus simulation
Aspen = win32.Dispatch('Apwn.document')
Aspen.InitFromArchive2(os.path.abspath(Aspen_file_path))
Aspen.Engine.Run2()
Aspen.Visible = True # Optional
Aspen.SuppressDialogs = 1

# Defining flue gas composition
flue_gas_feed_flowrate = float(df_flue_gas.iloc[0].iloc[1]) # t/h
composition_flue_gas_N2_in = float(df_flue_gas.iloc[0].iloc[2] / 100) # wt.% 
composition_flue_gas_O2_in = float(df_flue_gas.iloc[0].iloc[3] / 100) # wt.%
composition_flue_gas_CO2_in = float(df_flue_gas.iloc[0].iloc[4] / 100) # wt.%
composition_flue_gas_H2O_in = float(df_flue_gas.iloc[0].iloc[5] / 100) # wt.%
composition_flue_gas_H2_in = float(df_flue_gas.iloc[0].iloc[6] / 100) # wt.%
composition_flue_gas_CO_in = float(df_flue_gas.iloc[0].iloc[7] / 100) # wt.%
composition_flue_gas_CH4_in = float(df_flue_gas.iloc[0].iloc[8] / 100) # wt.%

# Update data for Flue gas stream in the simulation
Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Input\TOTFLOW\MIXED').Value = flue_gas_feed_flowrate
Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Input\FLOW\MIXED\N2').Value = composition_flue_gas_N2_in
Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Input\FLOW\MIXED\O2').Value = composition_flue_gas_O2_in
Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Input\FLOW\MIXED\CO2').Value = composition_flue_gas_CO2_in
Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Input\FLOW\MIXED\H2O').Value = composition_flue_gas_H2O_in
#Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Input\FLOW\MIXED\H2').Value = composition_flue_gas_H2_in
#Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Input\FLOW\MIXED\CO').Value = composition_flue_gas_CO_in
#Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Input\FLOW\MIXED\CH4').Value = composition_flue_gas_CH4_in

# Defining lean solvent data
lean_loading = 0.12 # mol NH3/mol CO2 [0.1-0.2]
molecular_weight_NH3 = 17 # g/mol
molecular_weight_CO2 = 44 # g/mol
composition_lean_solvent_NH3_in = 0.05 # wt.%
composition_lean_solvent_CO2_in = ((composition_lean_solvent_NH3_in) / ((molecular_weight_NH3 / molecular_weight_CO2) * (1 / lean_loading))) # wt.%
composition_lean_solvent_H2O_in = 1 - composition_lean_solvent_NH3_in - composition_lean_solvent_CO2_in # wt.%
minimum_solvent_flowrate = 1 * flue_gas_feed_flowrate # t/h
maximum_solvent_flowrate = 3.5 * flue_gas_feed_flowrate # t/h

# Update data for Lean solvent stream in the simulation
Aspen.Tree.FindNode(r'\Data\Streams\LEANNH3\Input\FLOW\MIXED\NH3').Value = composition_lean_solvent_NH3_in
Aspen.Tree.FindNode(r'\Data\Streams\LEANNH3\Input\FLOW\MIXED\CO2').Value = composition_lean_solvent_CO2_in
Aspen.Tree.FindNode(r'\Data\Streams\LEANNH3\Input\FLOW\MIXED\H2O').Value = composition_lean_solvent_H2O_in

# Define the range for solvent flowrate
solvent_flowrates_array = np.linspace(minimum_solvent_flowrate, maximum_solvent_flowrate, 100)

# ----- AsPyCC: Absorber Design ------

# Lists to store results for converged and non-converged cases
ccr_converged, ccr_not_converged = [], []
solvent_flowrates_converged, solvent_flowrates_not_converged = [], []
Aspen.Reinit()

# Loop over solvent flowrates to find CCR and store results
for flowrate in solvent_flowrates_array:
    Aspen.Tree.FindNode(r'\Data\Streams\LEANNH3\Input\TOTFLOW\MIXED').Value = flowrate
    Aspen.Engine.Run2()

    # Wait for Aspen to complete calculations
    while Aspen.Engine.IsRunning:
        time.sleep(0.5)
    
    try:
        status_message = Aspen.Tree.FindNode(r'\Data\Results Summary\Run-Status\Output\PER_ERROR').Value
        if status_message == 0:
            
            # Calculate CCR if no errors in status message
            clean_gas_CO2_out = Aspen.Tree.FindNode(r'\Data\Streams\CLEANGAS\Output\MASSFLOW\MIXED\CO2').Value
            flue_gas_CO2_in = Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Output\MASSFLOW\MIXED\CO2').Value
            ccr = ((flue_gas_CO2_in - clean_gas_CO2_out) / flue_gas_CO2_in) * 100
            ccr_converged.append(ccr)
            solvent_flowrates_converged.append(float(flowrate))
            if 89.00 <= ccr <= 90.99:
                print(f'CCR target reached: {ccr:.2f}')
                break
        else:
            clean_gas_CO2_out = Aspen.Tree.FindNode(r'\Data\Streams\CLEANGAS\Output\MASSFLOW\MIXED\CO2').Value
            ccr_not_converged.append(clean_gas_CO2_out)
            solvent_flowrates_not_converged.append(float(flowrate))
    
    except AttributeError:
        break

# Display results for converged and non-converged simulations (Optional)
#print(f'Successful simulations: {len(ccr_converged)}')
#print(f'Unsuccessful simulations: {len(ccr_not_converged)}')
#print(f'Infinite height column CCR: {ccr_converged[-1]:.2f} %')
#print(f'Minimum solvent flowrate: {solvent_flowrates_converged[-1]:.2f} t/h')

# Update to effective solvent flowrate and re-run simulation
solvent_factor = 1.1  # Adjustable factor for solvent flowrate
effective_solvent_flowrate = solvent_flowrates_converged[-1] * solvent_factor
Aspen.Tree.FindNode(r'\Data\Streams\LEANNH3\Input\TOTFLOW\MIXED').Value = round(effective_solvent_flowrate, 2)
Aspen.Engine.Run2()
while Aspen.Engine.IsRunning:
    time.sleep(0.5)
clean_gas_CO2_out = Aspen.Tree.FindNode(r'\Data\Streams\CLEANGAS\Output\MASSFLOW\MIXED\CO2').Value
flue_gas_CO2_in = Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Output\MASSFLOW\MIXED\CO2').Value
actual_ccr = ((flue_gas_CO2_in - clean_gas_CO2_out) / flue_gas_CO2_in) * 100

# Simulate CCR at different column heights
heights = np.linspace(100, 5, 100)
ccr_at_different_heights, heights_converged = [], []
for height in heights:
    Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Subobjects\Column Internals\INT-1\Input\CA_PACK_HT\INT-1\CS-1').Value = height
    Aspen.Engine.Run2()
    while Aspen.Engine.IsRunning:
        time.sleep(0.5)
    try:
        status_message = Aspen.Tree.FindNode(r'\Data\Results Summary\Run-Status\Output\PER_ERROR').Value
        if status_message == 0:
            clean_gas_CO2_out = Aspen.Tree.FindNode(r'\Data\Streams\CLEANGAS\Output\MASSFLOW\MIXED\CO2').Value
            flue_gas_CO2_in = Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Output\MASSFLOW\MIXED\CO2').Value
            ccr = ((flue_gas_CO2_in - clean_gas_CO2_out) / flue_gas_CO2_in) * 100
            ccr_at_different_heights.append(ccr)
            heights_converged.append(height)
            if 89.00 <= ccr <= 90.99:
                print(f'CCR target reached at height {height:.2f} m')
                break
    except AttributeError:
        break
    
# Adjust column diameter and solvent flowrate to meet flooding and CCR targets
diameter = Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Subobjects\Column Internals\INT-1\Input\CA_DIAM\INT-1\CS-1').Value
current_solvent_flowrate = Aspen.Tree.FindNode(r'\Data\Streams\LEANNH3\Input\TOTFLOW\MIXED').Value
flooding_limits = (69.99, 79.99)
ccr_limits = (84.90, 90.99)
counter, max_iterations = 1, 100
diameter_list, ccr_list, solvent_list, flooding_list = [], [], [], []
while counter <= max_iterations:
    Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Subobjects\Column Internals\INT-1\Input\CA_DIAM\INT-1\CS-1').Value = diameter
    Aspen.Tree.FindNode(r'\Data\Streams\LEANNH3\Input\TOTFLOW\MIXED').Value = current_solvent_flowrate
    Aspen.Engine.Run2()
    while Aspen.Engine.IsRunning:
        time.sleep(0.5)
    current_flooding = Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Output\CA_FLD_FAC1\INT-1\CS-1').Value
    clean_gas_CO2_out = Aspen.Tree.FindNode(r'\Data\Streams\CLEANGAS\Output\MASSFLOW\MIXED\CO2').Value
    flue_gas_CO2_in = Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Output\MASSFLOW\MIXED\CO2').Value
    current_ccr = ((flue_gas_CO2_in - clean_gas_CO2_out) / flue_gas_CO2_in) * 100

    # Store results for each iteration
    diameter_list.append(diameter)
    ccr_list.append(current_ccr)
    solvent_list.append(current_solvent_flowrate)
    flooding_list.append(current_flooding)
    if flooding_limits[0] < current_flooding < flooding_limits[1] and ccr_limits[0] < current_ccr < ccr_limits[1]:
        break
    if current_flooding > flooding_limits[1]:
        diameter += 0.25
    elif current_flooding < flooding_limits[0]:
        diameter -= 0.25
    if current_ccr > ccr_limits[1]:
        current_solvent_flowrate -= 1
    elif current_ccr < ccr_limits[0]:
        current_solvent_flowrate += 1
    counter += 1

# Adjust column height for final CCR range
current_height = Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Subobjects\Column Internals\INT-1\Input\CA_PACK_HT\INT-1\CS-1').Value
ccr_limits_final = (89.00, 90.99)
counter = 1
while counter <= max_iterations:
    Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Subobjects\Column Internals\INT-1\Input\CA_PACK_HT\INT-1\CS-1').Value = current_height
    Aspen.Engine.Run2()
    while Aspen.Engine.IsRunning:
        time.sleep(0.5)
    current_ccr = ((flue_gas_CO2_in - clean_gas_CO2_out) / flue_gas_CO2_in) * 100
    if ccr_limits_final[0] < current_ccr < ccr_limits_final[1]:
        break
    if current_ccr > ccr_limits_final[1]:
        current_height -= 0.1
    elif current_ccr < ccr_limits_final[0]:
        current_height += 0.1
    counter += 1
    
# Final results
final_column_height = Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Subobjects\Column Internals\INT-1\Input\CA_PACK_HT\INT-1\CS-1').Value
final_column_diameter = Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Subobjects\Column Internals\INT-1\Input\CA_DIAM\INT-1\CS-1').Value
final_flooding = Aspen.Tree.FindNode(r'\Data\Blocks\ABSORBER\Output\CA_FLD_FAC1\INT-1\CS-1').Value
final_solvent_flowrate = Aspen.Tree.FindNode(r'\Data\Streams\LEANNH3\Input\TOTFLOW\MIXED').Value
clean_gas_CO2_out = Aspen.Tree.FindNode(r'\Data\Streams\CLEANGAS\Output\MASSFLOW\MIXED\CO2').Value
flue_gas_CO2_in = Aspen.Tree.FindNode(r'\Data\Streams\FLUEGAS\Output\MASSFLOW\MIXED\CO2').Value
final_ccr = ((flue_gas_CO2_in - clean_gas_CO2_out) / flue_gas_CO2_in) * 100

# ----- AsPyCC: Heat exchanger and stripper design -----

# Create heat exchenger prior stripper
Aspen.Tree.Elements('Data').Elements('Blocks').Elements.Add('HXT1' + '!' + 'Heater')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('HXT1').Elements('Ports').Elements('F(IN)').Elements.Add('RICHSOLV')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('TOSTRIP' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('HXT1').Elements('Ports').Elements('P(OUT)').Elements.Add('TOSTRIP')
Aspen.Tree.FindNode(r'\Data\Blocks\HXT1\Input\TEMP').Value = 135 # °C
Aspen.Tree.FindNode(r'\Data\Blocks\HXT1\Input\PRES').Value = 5 # bar

# Run simulation
Aspen.Engine.Run2()
while Aspen.Engine.IsRunning:
        time.sleep(0.5)

# Add Stripper and condenser with the streams Vapor, Reflux, CO2, Leansolv
Aspen.Tree.Elements('Data').Elements('Blocks').Elements.Add('STRIP' + '!' + 'RadFrac')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('STRIP').Elements('Ports').Elements('F(IN)').Elements.Add('TOSTRIP')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('LEANSOLV' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('STRIP').Elements('Ports').Elements('B(OUT)').Elements.Add('LEANSOLV')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements.Add('CNDNSR' + '!' + 'Flash2')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('VAPOR' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('STRIP').Elements('Ports').Elements('VD(OUT)').Elements.Add('VAPOR')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('CNDNSR').Elements('Ports').Elements('F(IN)').Elements.Add('VAPOR')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('CO2' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('REFLUX' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('CNDNSR').Elements('Ports').Elements('V(OUT)').Elements.Add('CO2')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('CNDNSR').Elements('Ports').Elements('L(OUT)').Elements.Add('REFLUX')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('STRIP').Elements('Ports').Elements('F(IN)').Elements.Add('REFLUX')

# Stripper stages
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\NSTAGE').Value = 10

# Stripper condenser
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\CONDENSER').Value = 'NONE'

# Stripper boil-up ratio
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\BASIS_BR').Value = 0.03

# Feed stages
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\FEED_STAGE\TOSTRIP').Value = 1
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\FEED_STAGE\REFLUX').Value = 1

# Convention of stages
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\FEED_CONVE2\TOSTRIP').Value = 'ABOVE-STAGE'
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\FEED_CONVE2\REFLUX').Value = 'ABOVE-STAGE'

# Stripper pressure
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\PRES1').Value = 5 

# Flash temperature and pressure
Aspen.Tree.FindNode(r'\Data\Blocks\CNDNSR\Input\TEMP').Value = 30
Aspen.Tree.FindNode(r'\Data\Blocks\CNDNSR\Input\PRES').Value = 0

# Run simulation
Aspen.Engine.Run2()
while Aspen.Engine.IsRunning:
        time.sleep(0.5)
        
# ----- Cross-heat exchanger integration -----

# Cross-heat exchanger
Aspen.Tree.Elements('Data').Elements('Blocks').Elements.Add('CHXT' + '!' + 'MHeatX')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('CHXT').Elements('Ports').Elements('HF(IN)').Elements.Add('LEANSOLV')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('HXT1').Elements('Ports').Elements('F(IN)').Elements.Remove('RICHSOLV')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('CHXT').Elements('Ports').Elements('CF(IN)').Elements.Add('RICHSOLV')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('HOTRICH' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('HXT1').Elements('Ports').Elements('F(IN)').Elements.Add('HOTRICH')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('COLDLEAN' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('CHXT').Elements('Ports').Elements('HP(OUT)').Elements.Add('COLDLEAN')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('CHXT').Elements('Ports').Elements('CP(OUT)').Elements.Add('HOTRICH')

# Hot outlet stream
Aspen.Tree.FindNode(r'\Data\Blocks\CHXT\Input\OUT\LEANSOLV').Value = 'COLDLEAN'

# Cold outlet stream
Aspen.Tree.FindNode(r'\Data\Blocks\CHXT\Input\OUT\RICHSOLV').Value = 'HOTRICH'

# Cross-heat exchanger temperature
Aspen.Tree.FindNode(r'\Data\Blocks\CHXT\Input\SPEC\LEANSOLV').Value = 'TEMP'
Aspen.Tree.FindNode(r'Data\Blocks\CHXT\Input\VALUE\LEANSOLV').Value = 50

# Run simulation
Aspen.Engine.Run2()
while Aspen.Engine.IsRunning:
    time.sleep(0.5)
    
# -----Recycle-loading correction -----

# Compute the ammount of MEA for the make-up
make_up_flowrate = Aspen.Tree.FindNode(r'\Data\Streams\CLEANGAS\Output\MOLEFLOW\MIXED\NH3').Value + Aspen.Tree.FindNode(r'\Data\Streams\CO2\Output\MOLEFLOW\MIXED\NH3').Value

# Create and setup the make-up stream
Aspen.Tree.Elements('Data').Elements('Blocks').Elements.Add('MIXER' + '!' + 'Mixer')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('MKP' + '!' + 'MATERIAL')
Aspen.Tree.FindNode(r'\Data\Streams\MKP\Input\TEMP\MIXED').Value = 15
Aspen.Tree.FindNode(r'\Data\Streams\MKP\Input\PRES\MIXED').Value = 1
Aspen.Tree.FindNode(r'\Data\Streams\MKP\Input\TOTFLOW\MIXED').Value = make_up_flowrate
Aspen.Tree.FindNode(r'\Data\Streams\MKP\Input\FLOW\MIXED\NH3').Value = make_up_flowrate
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('MIXER').Elements('Ports').Elements('F(IN)').Elements.Add('COLDLEAN')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('MIXER').Elements('Ports').Elements('F(IN)').Elements.Add('MKP')

# Create cooler and setup streams
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('TOCOOLER' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('MIXER').Elements('Ports').Elements('P(OUT)').Elements.Add('TOCOOLER')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements.Add('HXT2' + '!' + 'Heater')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('HXT2').Elements('Ports').Elements('F(IN)').Elements.Add('TOCOOLER')
Aspen.Tree.Elements('Data').Elements('Streams').Elements.Add('RECYCLE' + '!' + 'MATERIAL')
Aspen.Tree.Elements('Data').Elements('Blocks').Elements('HXT2').Elements('Ports').Elements('P(OUT)').Elements.Add('RECYCLE')
Aspen.Tree.FindNode(r'\Data\Blocks\HXT2\Input\TEMP').Value = 15 # °C
Aspen.Tree.FindNode(r'\Data\Blocks\HXT2\Input\PRES').Value = 1 # bar
Aspen.Engine.Run2()
while Aspen.Engine.IsRunning:
    time.sleep(0.5)

# Computing the apparent lean loading
composition_lean_rec_out_NH3 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\NH3').Value
composition_lean_rec_out_NH4 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\NH4+').Value
composition_lean_rec_out_NH2COO = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\NH2COO-').Value
composition_lean_rec_out_CO2 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\CO2').Value
composition_lean_rec_out_HCO3 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\HCO3-').Value
composition_lean_rec_out_CO3 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\CO3-2').Value
apparent_CO2 = (composition_lean_rec_out_CO2 + composition_lean_rec_out_HCO3 + composition_lean_rec_out_CO3 + composition_lean_rec_out_NH2COO)
apparent_NH3 = (composition_lean_rec_out_NH3 + composition_lean_rec_out_NH4 + composition_lean_rec_out_NH2COO)
apparent_lean_loading = apparent_CO2 / apparent_NH3

# Initial boil-up ratio
boilup_ratio = 0.03
step = 0.001
tolerance = 0.001
while True:
    
    # Set the boil-up ratio in Aspen
    Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\BASIS_BR').Value = boilup_ratio
    
    # Run the simulation
    Aspen.Engine.Run2()
    while Aspen.Engine.IsRunning:
        time.sleep(0.5)
    
    # Compute the new lean loading
    # Computing the apparent lean loading
    composition_lean_rec_out_NH3 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\NH3').Value
    composition_lean_rec_out_NH4 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\NH4+').Value
    composition_lean_rec_out_NH2COO = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\NH2COO-').Value
    composition_lean_rec_out_CO2 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\CO2').Value
    composition_lean_rec_out_HCO3 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\HCO3-').Value
    composition_lean_rec_out_CO3 = Aspen.Tree.FindNode(r'\Data\Streams\RECYCLE\Output\MOLEFRAC\MIXED\CO3-2').Value
    apparent_CO2 = (composition_lean_rec_out_CO2 + composition_lean_rec_out_HCO3 + composition_lean_rec_out_CO3 + composition_lean_rec_out_NH2COO)
    apparent_NH3 = (composition_lean_rec_out_NH3 + composition_lean_rec_out_NH4 + composition_lean_rec_out_NH2COO)
    calculated_loading = apparent_CO2 / apparent_NH3
    
    # Check if the loading is within the tolerance
    if abs(calculated_loading - lean_loading) <= tolerance:
        break  # Converged

    # Adjust the boil-up ratio
    if calculated_loading > lean_loading :
        boilup_ratio += step
    else:
        boilup_ratio -= step
        
# Add utilities
Aspen.Tree.FindNode(r'\Data\Blocks\HXT1\Input\UTILITY_ID').Value = 'U-2'
Aspen.Tree.FindNode(r'\Data\Blocks\HXT2\Input\UTILITY_ID').Value = 'U-4'
Aspen.Tree.FindNode(r'\Data\Blocks\STRIP\Input\REB_UTIL').Value = 'U-2'
Aspen.Tree.FindNode(r'\Data\Blocks\CNDNSR\Input\UTILITY_ID').Value = 'U-3'
Aspen.Engine.Run2()
while Aspen.Engine.IsRunning:
    time.sleep(0.5)
    
# At this point the economics are activated in the simulation file, and the results can be retrieved.
# Finally, the simulation file is closed. It is recommended to save the simulation as a new compound file, as the .bkp file will be used for future simulations.

# Close the COM connection
Aspen.Close()