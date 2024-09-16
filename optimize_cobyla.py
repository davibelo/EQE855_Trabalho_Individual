import os
import win32com.client as win32
from scipy.optimize import minimize

file = r"RECAP_revK.bkp"
aspen_Path = os.path.abspath(file)

print('Connecting to the Aspen Plus... Please wait ')
Application = win32.Dispatch('Apwn.Document')  # Registered name of Aspen Plus
print('Connected!')

Application.InitFromArchive2(aspen_Path)
Application.visible = 0

# Initial guess
x0 = [560000, 950000, 3]

def simulate(x):
    QN1, QN2, QC = x
    Application.Tree.FindNode(r"\Data\Blocks\N-640\Input\QN").Value = QN1
    Application.Tree.FindNode(r"\Data\Blocks\N-641\Input\QN").Value = QN2
    Application.Tree.FindNode(r"\Data\Blocks\N-641\Input\Q1").Value = QC
    Application.Engine.Run2()
    cH2S = Application.Tree.FindNode(r"\Data\Streams\AGUAR1\Output\MOLEFRAC\MIXED\H2S").Value
    cNH3 = Application.Tree.FindNode(r"\Data\Streams\AGUAR1\Output\MOLEFRAC\MIXED\NH3").Value
    cH2S_ppm = cH2S * 1E6
    cNH3_ppm = cNH3 * 1E6
    print(f"Simulating with QN1: {round(QN1, 0)}, QN2: {round(QN2, 0)}, QC: {round(QC, 2)} -> H2S: {round(cH2S_ppm, 3)}, NH3: {round(cNH3_ppm, 3)}")
    return cH2S_ppm, cNH3_ppm

# Objective function to minimize (without scaling)
def cost(x):
    QN1, QN2, QC = x
    total_cost = QN1 + QN2 + QC
    return total_cost

# Constraint 1 (H2S PPM <= 0.2)
def constraint1(x):
    cH2S_ppm, _ = simulate(x)
    return 0.2 - cH2S_ppm  # >= 0

# Constraint 2 (NH3 PPM <= 15)
def constraint2(x):
    _, cNH3_ppm = simulate(x)
    return 15 - cNH3_ppm  # >= 0

# Lower and upper bound constraints for QN1, QN2, and QC
def bound_QN1_lower(x):
    QN1_lower = 450000
    return x[0] - QN1_lower

def bound_QN1_upper(x):
    QN1_upper = 600000
    return QN1_upper - x[0]

def bound_QN2_lower(x):
    QN2_lower = 700000
    return x[1] - QN2_lower

def bound_QN2_upper(x):
    QN2_upper = 1200000
    return QN2_upper - x[1]

def bound_QC_lower(x):
    QC_lower = 1
    return x[2] - QC_lower

def bound_QC_upper(x):
    QC_upper = 5
    return QC_upper - x[2]

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

options = {
    'maxiter': 10000,  # Increase iterations if necessary
    'tol': 1e-3       # Loosen the constraint satisfaction tolerance (increase this value for more leniency)
}

# Solving the optimization problem with COBYLA
result = minimize(cost, x0, method='COBYLA', constraints=constraints, options=options)

# Output results and check maxcv (maximum constraint violation)
opt = result.x
cost_min = result.fun
num_function_evals = result.nfev
success = result.success
message = result.message
maxcv = result.maxcv  # Magnitude of constraint violation

print('Optimal values: ', opt)
print('Minimum cost: ', cost_min)
print('Number of function evaluations: ', num_function_evals)
print('Optimization success: ', success)
print('Message: ', message)
print('Maximum constraint violation (maxcv): ', maxcv)
