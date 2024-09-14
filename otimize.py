import os
import win32com.client as win32
from scipy.optimize import minimize

file = r"RECAP_revK.bkp"
aspen_Path = os.path.abspath(file)

print('Connecting to the Aspen Plus... Please wait ')
Application = win32.Dispatch('Apwn.Document') # Registered name of Aspen Plus
print('Connected!')

Application.InitFromArchive2(aspen_Path)
Application.visible = 0

def simulate(x):
    QN1, QN2, QC = x
    Application.Tree.FindNode(r"\Data\Blocks\N-640\Input\QN").Value = QN1
    Application.Tree.FindNode(r"\Data\Blocks\N-641\Input\QN").Value = QN2
    Application.Tree.FindNode(r"\Data\Blocks\N-641\Input\Q1").Value = QC
    Application.Engine.Run2()
    cH2S = Application.Tree.FindNode(r"\Data\Streams\AGUAR1\Output\MOLEFRAC\MIXED\H2S").Value
    cNH3 = Application.Tree.FindNode(r"\Data\Streams\AGUAR1\Output\MOLEFRAC\MIXED\NH3").Value
    cH2S_ppm = cH2S*1E6
    cNH3_ppm = cNH3*1E6
    y = cH2S_ppm, cNH3_ppm
    print(f"Simulating with QN1: {QN1}, QN2: {QN2}, QC: {QC} -> H2S: {cH2S_ppm}, NH3: {cNH3_ppm}")
    total_cost = QN1 + QN2 + QC
    print(f"Total Cost: {total_cost}")
    return y

# Objective function to minimize
def cost(x):
    QN1, QN2, QC = x
    total_cost = QN1 + QN2 + QC
    return total_cost

# Constraint 1
def constraint1(x):
    cH2S_ppm, _ = simulate(x)
    return cH2S_ppm - 0.2 # >= 0.2ppm
    
# Constraint 2
def constraint2(x):
    _, cNH3_ppm = simulate(x)
    return cNH3_ppm - 15  # >= 15ppm

# Initial guess
x0 = [560000, 950000, 3]

# Constraints as a dictionary
constraints = [{'type': 'ineq', 'fun': constraint1}]
constraints = [{'type': 'ineq', 'fun': constraint1}]

# Bounds
bounds = [(520000, 600000), (900000, 1200000), (1,5)]

# Solving the optimization problem
#result = minimize(cost, x0, method='trust-constr', bounds=bounds, constraints=constraints)
result = minimize(cost, x0, method='SLSQP', bounds=bounds, constraints=constraints)

# Output results
opt = result.x
cost_min = result.fun
print('Optimal values: ', opt)
