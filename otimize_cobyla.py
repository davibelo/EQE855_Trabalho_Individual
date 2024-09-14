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

# Initial guess
x0 = [560000, 950000, 3]

# Scaling factors
scale_factors = [1e5, 1e5, 1]  # Scaling for QN1, QN2, QC

def simulate(x_scaled):
    x = [x_scaled[0] * scale_factors[0], x_scaled[1] * scale_factors[1], x_scaled[2] * scale_factors[2]]
    QN1, QN2, QC = x
    Application.Tree.FindNode(r"\Data\Blocks\N-640\Input\QN").Value = QN1
    Application.Tree.FindNode(r"\Data\Blocks\N-641\Input\QN").Value = QN2
    Application.Tree.FindNode(r"\Data\Blocks\N-641\Input\Q1").Value = QC
    Application.Engine.Run2()
    cH2S = Application.Tree.FindNode(r"\Data\Streams\AGUAR1\Output\MOLEFRAC\MIXED\H2S").Value
    cNH3 = Application.Tree.FindNode(r"\Data\Streams\AGUAR1\Output\MOLEFRAC\MIXED\NH3").Value
    cH2S_ppm = cH2S * 1E6
    cNH3_ppm = cNH3 * 1E6
    y = cH2S_ppm, cNH3_ppm
    print(f"Simulating with QN1: {QN1}, QN2: {QN2}, QC: {QC} -> H2S: {cH2S_ppm}, NH3: {cNH3_ppm}")
    return y

# Objective function to minimize (with scaling)
def cost(x_scaled):
    x = [x_scaled[0] * scale_factors[0], x_scaled[1] * scale_factors[1], x_scaled[2] * scale_factors[2]]
    total_cost = x[0] + x[1] + x[2]
    return total_cost

# Constraint 1 (H2S PPM >= 0.2)
def constraint1(x_scaled):
    cH2S_ppm, _ = simulate(x_scaled)
    return cH2S_ppm - 0.2  # >= 0.2 ppm

# Constraint 2 (NH3 PPM >= 15)
def constraint2(x_scaled):
    _, cNH3_ppm = simulate(x_scaled)
    return cNH3_ppm - 15  # >= 15 ppm

# Lower bound constraint for QN1
def bound_QN1_lower(x_scaled):
    QN1_lower = 510000
    return x_scaled[0] - (QN1_lower / scale_factors[0])

# Upper bound constraint for QN1
def bound_QN1_upper(x_scaled):
    QN1_upper = 600000
    return (QN1_upper / scale_factors[0]) - x_scaled[0]

# Lower bound constraint for QN2
def bound_QN2_lower(x_scaled):
    QN2_lower = 900000
    return x_scaled[1] - (QN2_lower / scale_factors[1])

# Upper bound constraint for QN2
def bound_QN2_upper(x_scaled):
    QN2_upper = 1200000
    return (QN2_upper / scale_factors[1]) - x_scaled[1]

# Lower bound constraint for QC
def bound_QC_lower(x_scaled):
    QC_lower = 1
    return x_scaled[2] - (QC_lower / scale_factors[2])

# Upper bound constraint for QC
def bound_QC_upper(x_scaled):
    QC_upper = 5
    return (QC_upper / scale_factors[2]) - x_scaled[2]

# Initial guess (with scaling)
x0_scaled = [x0[i] / scale_factors[i] for i in range(3)]

# Define constraints as a list of dictionaries
constraints = [
    {'type': 'ineq', 'fun': constraint1},      # H2S constraint
    {'type': 'ineq', 'fun': constraint2},      # NH3 constraint
    {'type': 'ineq', 'fun': bound_QN1_lower},  # QN1 lower bound
    {'type': 'ineq', 'fun': bound_QN1_upper},  # QN1 upper bound
    {'type': 'ineq', 'fun': bound_QN2_lower},  # QN2 lower bound
    {'type': 'ineq', 'fun': bound_QN2_upper},  # QN2 upper bound
    {'type': 'ineq', 'fun': bound_QC_lower},   # QC lower bound
    {'type': 'ineq', 'fun': bound_QC_upper}    # QC upper bound
]

# Solving the optimization problem with COBYLA
result = minimize(cost, x0_scaled, method='COBYLA', constraints=constraints, options={'disp': True})

# Rescale the results
opt_scaled = result.x
opt = [opt_scaled[i] * scale_factors[i] for i in range(3)]

# Output results
print('Optimal values: ', opt)
