import os
import numpy as np
import win32com.client as win32
from scipy.optimize import minimize
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

file = r"RECAP_revK.bkp"
aspen_Path = os.path.abspath(file)

print('Connecting to the Aspen Plus... Please wait ')
Application = win32.Dispatch('Apwn.Document') # Registered name of Aspen Plus
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

# Bounds
bounds = [(450000, 600000), (700000, 1200000), (1,5)]

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
constraints = [{'type': 'ineq', 'fun': constraint1}, {'type': 'ineq', 'fun': constraint2}]

# Solving the optimization problem
result = minimize(cost, x0_scaled, method='SLSQP', bounds=bounds_scaled, constraints=constraints, options={'ftol': 1e-6})
# result = minimize(cost, x0_scaled, method='trust-constr', bounds=bounds_scaled, constraints=constraints)

# Rescale the results
opt_scaled = result.x
opt = [opt_scaled[i] * scale_factors[i] for i in range(3)]

# Output results
cost_min = result.fun
num_iterations = result.nit
num_function_evals = result.nfev

log_message(f'Optimal values: {opt}')
log_message(f'Minimum cost: {cost_min}')
log_message(f'Number of iterations: {num_iterations}')
log_message(f'Number of function evaluations: {num_function_evals}')

# Final simulation with optimal values
simulate(opt_scaled, print_temperature=True)

# Close log file
log_file.close()

# Convert x_values to a format that can be plotted (split QN1, QN2, QC)
QN1_values = [x[0] for x in x_values]
QN2_values = [x[1] for x in x_values]
QC_values = [x[2] for x in x_values]

# Create a single figure with 3 subplots
fig = plt.figure(figsize=(18, 6))

# Custom angles for the plots
elev_angle = 30 # Elevation angle (default is 30 degrees)
azim_angle = 110 # Azimuth angle (default is 120 degrees)

# Padding value between axis tick values and axis titles
axis_labelpad = 10

# Plot 1: QN1 vs QN2 vs Objective Function
ax1 = fig.add_subplot(131, projection='3d')
ax1.plot(QN1_values, QN2_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax1.scatter(opt[0], opt[1], cost_min, color='blue', s=100, label='Optimal Point')
ax1.set_xlabel('QN1', labelpad=axis_labelpad)
ax1.set_ylabel('QN2', labelpad=axis_labelpad)
ax1.set_zlabel('Cost', labelpad=axis_labelpad)
ax1.set_title('QN1 vs QN2 vs Cost', pad=20)  # Add padding to title
ax1.view_init(elev=elev_angle, azim=azim_angle)  # Set custom view angles
ax1.legend()

# Plot 2: QN1 vs QC vs Objective Function
ax2 = fig.add_subplot(132, projection='3d')
ax2.plot(QN1_values, QC_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax2.scatter(opt[0], opt[2], cost_min, color='blue', s=100, label='Optimal Point')
ax2.set_xlabel('QN1', labelpad=axis_labelpad)
ax2.set_ylabel('QC', labelpad=axis_labelpad)
ax2.set_zlabel('Cost', labelpad=axis_labelpad)
ax2.set_title('QN1 vs QC vs Cost', pad=20)  # Add padding to title
ax2.view_init(elev=elev_angle, azim=azim_angle)  # Set custom view angles
ax2.legend()

# Plot 3: QN2 vs QC vs Objective Function
ax3 = fig.add_subplot(133, projection='3d')
ax3.plot(QN2_values, QC_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax3.scatter(opt[1], opt[2], cost_min, color='blue', s=100, label='Optimal Point')
ax3.set_xlabel('QN2', labelpad=axis_labelpad)
ax3.set_ylabel('QC', labelpad=axis_labelpad)
ax3.set_zlabel('Cost', labelpad=axis_labelpad)
ax3.set_title('QN2 vs QC vs Cost', pad=20)  # Add padding to title
ax3.view_init(elev=elev_angle, azim=azim_angle)  # Set custom view angles
ax3.legend()

# Adjust the overall layout with margins to avoid trimming
plt.subplots_adjust(left=0.05, right=0.95, top=0.90, bottom=0.10, wspace=0.3)

# Save the combined figure with the script name
figure_name = script_name + '_3d_plots.png'
figure_path = os.path.join(os.getcwd(), figure_name)
plt.savefig(figure_path, bbox_inches='tight')  # Ensure everything fits in the saved image

# Show the figure (optional)
plt.show()

print(f'3D plots saved as: {figure_path}')