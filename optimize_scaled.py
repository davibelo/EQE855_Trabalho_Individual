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

# Bounds
bounds = [(520000, 600000), (900000, 1200000), (1,5)]

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
    cH2S_ppm = cH2S*1E6
    cNH3_ppm = cNH3*1E6
    y = cH2S_ppm, cNH3_ppm
    print(f"Simulating with QN1: {round(QN1,2)}, QN2: {round(QN2,2)}, QC: {round(QC,2)} -> H2S: {round(cH2S_ppm,2)}, NH3: {round(cNH3_ppm,2)}")
    total_cost = QN1 + QN2 + QC
    print(f"Total Cost: {total_cost}")
    return y

# Objective function to minimize (with scaling)
def cost(x_scaled):
    x = [x_scaled[0] * scale_factors[0], x_scaled[1] * scale_factors[1], x_scaled[2] * scale_factors[2]]
    total_cost = x[0] + x[1] + x[2]
    return total_cost

# Constraint 1 (with scaling)
def constraint1(x_scaled):
    cH2S_ppm, _ = simulate(x_scaled)
    return 0.2 - cH2S_ppm  # <= 0.2ppm
    
# Constraint 2 (with scaling)
def constraint2(x_scaled):
    _, cNH3_ppm = simulate(x_scaled)
    return 15 - cNH3_ppm # <= 15ppm

# Initial guess (with scaling)
x0_scaled = [x0[i] / scale_factors[i] for i in range(3)]

# Bounds (with scaling)
bounds_scaled = [(low / scale_factors[i], high / scale_factors[i]) for i, (low, high) in enumerate(bounds)]

# Constraints as a dictionary
constraints = [{'type': 'ineq', 'fun': constraint1}]
constraints = [{'type': 'ineq', 'fun': constraint2}]

# Solving the optimization problem with SLSQP
#result = minimize(cost, x0_scaled, method='SLSQP', bounds=bounds_scaled, constraints=constraints, options={'ftol': 1e-5})
result = minimize(cost, x0_scaled, method='trust-constr', bounds=bounds_scaled, constraints=constraints)

# Rescale the results
opt_scaled = result.x
opt = [opt_scaled[i] * scale_factors[i] for i in range(3)]

# Output results

# Output results
opt = result.x
cost_min = result.fun
num_iterations = result.nit
num_function_evals = result.nfev

print('Optimal values: ', opt)
print('Minimum cost: ', cost_min)
print('Number of iterations: ', num_iterations)
print('Number of objective function evaluations: ', num_function_evals)
