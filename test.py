import os
import win32com.client as win32

# 1. Specify file name
file = r"RECAP_revJ\RECAP_revJ.bkp"

# 2. Get path to Aspen Plus file
aspen_Path = os.path.abspath(file)

# 3 Initiate Aspen Plus application
print('\n Connecting to the Aspen Plus... Please wait ')
Application = win32.Dispatch('Apwn.Document') # Registered name of Aspen Plus
print('Connected!')

# 4. Initiate Aspen Plus file
Application.InitFromArchive2(aspen_Path)

# 5. Make the files visible
Application.visible = 0


test_temp = Application.Tree.FindNode(r"\Data\Streams\G-ACID\Output\TEMP_OUT\MIXED").Value
print(f"temperature: {test_temp} °C")

test_mass_frac = Application.Tree.FindNode(r"\Data\Streams\AGUAPR1\Output\MASSFRAC\MIXED\H2S").Value
print(f"Mass Fraction H2S before: {test_mass_frac}")

test_reboiler_duty = Application.Tree.FindNode(r"\Data\Blocks\N-640\Input\QN").Value
print(f"Reboiler Duty before: {test_reboiler_duty}")

new_reboiler_duty = test_reboiler_duty + 10000
print(f"Modifying Reboiler Duty to {new_reboiler_duty}")
Application.Tree.FindNode(r"\Data\Blocks\N-640\Input\QN").Value = new_reboiler_duty

# Run the simulation after changing input variables
Application.Engine.Run2()

# Example: Reading output after simulation
test_mass_frac = Application.Tree.FindNode(r"\Data\Streams\AGUAPR1\Output\MASSFRAC\MIXED\H2S").Value
print(f"Mass Fraction H2S: {test_mass_frac}")

# Close the Aspen simulation
Application.Close()
