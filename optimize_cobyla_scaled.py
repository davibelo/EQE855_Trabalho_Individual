import os
import numpy as np
import win32com.client as win32
from scipy.optimize import minimize
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

file = r"RECAP_revK.bkp"
aspen_Path = os.path.abspath(file)

print('Connecting to the Aspen Plus... Please wait ')
Application = win32.Dispatch('Apwn.Document')  # Registered name of Aspen Plus
print('Connected!')

Application.InitFromArchive2(aspen_Path)
Application.visible = 0

# Create and open log file
script_name = os.path.splitext(os.path.basename(__file__))[0]
log_file_name = script_name + '.log'
log_file_path = os.path.join(os.getcwd(), log_file_name)
log_file = open(log_file_path, 'w')

# Initial guess
x0 = [560000, 950000, 3]

# Scaling factors
scale_factors = [1e5, 1e5, 1]  # Scaling for QN1, QN2, QC

# Lists to store non-scaled x values and corresponding objective function values
x_values = []
objective_values = []

def log_message(message):
    log_file.write(message + '\n')
    print(message)

def simulate(x_scaled, print_temperature: bool = False):
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
    message = f"Simulating with QN1: {round(QN1,0)}, QN2: {round(QN2,0)}, QC: {round(QC,2)} -> H2S: {round(cH2S_ppm,3)}, NH3: {round(cNH3_ppm,3)}"
    log_message(message)
    if print_temperature:
        T_bottom_N640 = Application.Tree.FindNode(r"\Data\Blocks\N-640\Output\B_TEMP\5").Value
        T_bottom_N641 = Application.Tree.FindNode(r"\Data\Blocks\N-641\Output\B_TEMP\6").Value
        T_top_N641 = Application.Tree.FindNode(r"\Data\Blocks\N-641\Output\B_TEMP\2").Value
        log_message(f"Temperatures: {T_bottom_N640}, {T_bottom_N641}, {T_top_N641}")
    return y

# Objective function to minimize (with scaling)
def cost(x_scaled):
    x = [x_scaled[0] * scale_factors[0], x_scaled[1] * scale_factors[1], x_scaled[2] * scale_factors[2]]
    total_cost = x[0] + x[1] + x[2]
    
    # Store the non-scaled x values and total cost
    x_values.append(x)  # Store non-scaled x
    objective_values.append(total_cost)  # Store objective function value
    
    return total_cost

# Constraint 1 (H2S PPM <= 0.2)
def constraint1(x_scaled):
    cH2S_ppm, _ = simulate(x_scaled)
    return 0.2 - cH2S_ppm  # >=0

# Constraint 2 (NH3 PPM <= 15)
def constraint2(x_scaled):
    _, cNH3_ppm = simulate(x_scaled)
    return 15 - cNH3_ppm  # >=0

# Lower bound constraint for QN1
def bound_QN1_lower(x_scaled):
    QN1_lower = 450000
    return x_scaled[0] - (QN1_lower / scale_factors[0])

# Upper bound constraint for QN1
def bound_QN1_upper(x_scaled):
    QN1_upper = 600000
    return (QN1_upper / scale_factors[0]) - x_scaled[0]

# Lower bound constraint for QN2
def bound_QN2_lower(x_scaled):
    QN2_lower = 700000
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

options = {
    'maxiter': 10000,
    'tol': 1e-2
}

# Solving the optimization problem with COBYLA
result = minimize(cost, x0_scaled, method='COBYLA', constraints=constraints, options=options)

# Rescale the results
opt_scaled = result.x
opt = [opt_scaled[i] * scale_factors[i] for i in range(3)]

# Output results and check maxcv (maximum constraint violation)
cost_min = result.fun
num_function_evals = result.nfev
success = result.success
message = result.message
maxcv = result.maxcv  # Magnitude of constraint violation

log_message(f'Optimal values: {opt}')
log_message(f'Minimum cost: {cost_min}')
log_message(f'Number of function evaluations: {num_function_evals}')
log_message(f'Optimization success: {success}')
log_message(f'Message: {message}')
log_message(f'Maximum constraint violation (maxcv): {maxcv}')

# Final simulation with optimal values
simulate(opt_scaled, print_temperature=True)

# Close log file
log_file.close()
# Plot the results (3D plot of QN1, QN2, and objective function)
fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')

# Convert x_values to a format that can be plotted (split QN1, QN2)
QN1_values = [x[0] for x in x_values]
QN2_values = [x[1] for x in x_values]
QC_values = [x[2] for x in x_values]

# Create the 3D scatter plot
ax.scatter(QN1_values, QN2_values, objective_values, c='r', marker='o')
ax.set_xlabel('QN1')
ax.set_ylabel('QN2')
ax.set_zlabel('Objective Function (Total Cost)')

# Save the figure with the script name
figure_name = script_name + '_3d_plot.png'
figure_path = os.path.join(os.getcwd(), figure_name)
plt.savefig(figure_path)

# Display the plot (optional)
plt.show()

print(f'3D plot saved as: {figure_path}')