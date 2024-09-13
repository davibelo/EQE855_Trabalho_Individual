import os
import win32com.client as win32

# 1. Specify file name
file = 'RECAP_revJ.bkp'

# 2. Get path to Aspen Plus file
aspen_Path = os.path.abspath(file)

# 3 Initiate Aspen Plus application
print('\n Connecting to the Aspen Plus... Please wait ')
Application = win32.Dispatch('Apwn.Document') # Registered name of Aspen Plus
print('Connected!')

# 4. Initiate Aspen Plus file
Application.InitFromArchive2(aspen_Path)

# 5. Make the files visible
Application.visible = 1

# Example: Reading input and output variables
test_temp = Application.Tree.FindNode(r"\Data\Streams\G-ACID\Output\TEMP_OUT\MIXED").Value
print(f"temperature: {test_temp} °C")

# Example: Reading output after simulation
test_mass_frac = Application.Tree.FindNode(r"\Data\Streams\AGUAPR1\Output\MASSFRAC\MIXED\H2S").Value
print(f"Mass Fraction H2S: {test_mass_frac}")

# Example: Modifying a variable
Application.Tree.FindNode(r"\Data\Blocks\N-640\Input\QN").Value = 565000

# Run the simulation after changing input variables
Application.Engine.Run2()

# Example: Reading output after simulation
test_mass_frac = Application.Tree.FindNode(r"\Data\Streams\AGUAPR1\Output\MASSFRAC\MIXED\H2S").Value
print(f"Mass Fraction H2S: {test_mass_frac}")

# Close the Aspen simulation
Application.Close()
