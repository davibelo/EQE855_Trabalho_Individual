import os
import win32com.client as win32

# Initialize Aspen Plus COM interface
aspen = win32.Dispatch('Apwn.Document')

# Load a specific Aspen Plus backup file (replace with your file path)
simulation_file = os.path.abspath('RECAP revJ copy.apwz')
aspen.InitFromArchive2(simulation_file)

# Example: Reading input and output variables
feed_temp = aspen.Tree.FindNode(r'\Data\Streams\G-ACID\Input\TEMP\MIXED').Value
print(f"temperature: {feed_temp} °C")

# Example: Modifying a variable
aspen.Tree.FindNode(r'\Data\Blocks\N-640\Input\QR').Value = 565000

# Run the simulation after changing input variables
aspen.Engine.Run2()

# Example: Reading output after simulation
mass_frac = aspen.Tree.FindNode("\Data\Streams\AGUAPR1\Output\MASSFRAC\MIXED\H2S")
print(f"Mass Fraction H2S: {mass_frac}")

# Close the Aspen simulation
aspen.Close()
