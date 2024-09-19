import os
import numpy as np
import win32com.client as win32
from scipy.optimize import minimize
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from scipy.interpolate import griddata

file = r"UTAA_revK.bkp"
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
x0 = [560000, 950000, 3, 0.5]

# Scaling factors
scale_factors = [1e5, 1e5, 1, 0.1]  # Scaling for QN1, QN2, QC

# Lists to store non-scaled x values and corresponding objective function values
x_values = []
objective_values = []

def log_message(message):
    log_file.write(message + '\n')
    print(message)

def simulate(x_scaled, print_temperature: bool = False):
    x = [x_scaled[0] * scale_factors[0], x_scaled[1] * scale_factors[1], x_scaled[2] * scale_factors[2], x_scaled[3] * scale_factors[3]]
    QN1, QN2, QC, SF = x    
    Application.Tree.FindNode(r"\Data\Blocks\T1\Input\QN").Value = QN1
    Application.Tree.FindNode(r"\Data\Blocks\T2\Input\QN").Value = QN2
    Application.Tree.FindNode(r"\Data\Blocks\T2\Input\Q1").Value = QC
    Application.Tree.FindNode(r"\Data\Blocks\SPLIT1\Input\FRAC\AGUAPR5A").Value = max(0, SF)
    Application.Engine.Run2()
    cH2S = Application.Tree.FindNode(r"\Data\Streams\AGUAR1\Output\MOLEFRAC\MIXED\H2S").Value
    cNH3 = Application.Tree.FindNode(r"\Data\Streams\AGUAR1\Output\MOLEFRAC\MIXED\NH3").Value
    cH2S_ppm = cH2S * 1E6
    cNH3_ppm = cNH3 * 1E6
    y = cH2S_ppm, cNH3_ppm
    message = f"Simulating with QN1: {round(QN1,0)}, QN2: {round(QN2,0)}, QC: {round(QC,2)}, SF: {round(SF,2)} -> H2S: {round(cH2S_ppm,3)}, NH3: {round(cNH3_ppm,3)}"
    log_message(message)
    if print_temperature:
        T_bottom_N640 = Application.Tree.FindNode(r"\Data\Blocks\T1\Output\B_TEMP\5").Value
        T_bottom_N641 = Application.Tree.FindNode(r"\Data\Blocks\T2\Output\B_TEMP\6").Value
        T_top_N641 = Application.Tree.FindNode(r"\Data\Blocks\T2\Output\B_TEMP\2").Value
        log_message(f"Temperatures: {T_bottom_N640}, {T_bottom_N641}, {T_top_N641}")
    return y

# Objective function to minimize (with scaling)
def cost(x_scaled):
    x = [x_scaled[0] * scale_factors[0], x_scaled[1] * scale_factors[1], x_scaled[2] * scale_factors[2], x_scaled[3] * scale_factors[3]]
    total_cost = x[0] + x[1] + x[2]
    # Store the non-scaled x values and total cost
    x_values.append(x)  # Store non-scaled x
    objective_values.append(total_cost)  # Store objective function value
    log_message(f"Total Cost: {total_cost}")
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

# Lower bound constraint for SF
def bound_SF_lower(x_scaled):
    SF_lower = 0
    return x_scaled[3] - (SF_lower / scale_factors[3])

# Upper bound constraint for SF
def bound_SF_upper(x_scaled):
    SF_upper = 1
    return (SF_upper / scale_factors[3]) - x_scaled[3]

# Initial guess (with scaling)
x0_scaled = [x0[i] / scale_factors[i] for i in range(4)]

# Define constraints as a list of dictionaries
constraints = [
    {'type': 'ineq', 'fun': constraint1},      # H2S constraint
    {'type': 'ineq', 'fun': constraint2},      # NH3 constraint
    {'type': 'ineq', 'fun': bound_QN1_lower},  # QN1 lower bound
    {'type': 'ineq', 'fun': bound_QN1_upper},  # QN1 upper bound
    {'type': 'ineq', 'fun': bound_QN2_lower},  # QN2 lower bound
    {'type': 'ineq', 'fun': bound_QN2_upper},  # QN2 upper bound
    {'type': 'ineq', 'fun': bound_QC_lower},   # QC lower bound
    {'type': 'ineq', 'fun': bound_QC_upper},   # QC upper bound
    {'type': 'ineq', 'fun': bound_SF_lower},   # SF lower bound
    {'type': 'ineq', 'fun': bound_SF_upper}    # SF upper bound
]

options = {
    'maxiter': 10000,
    'tol': 1e-2
}

# Solving the optimization problem with COBYLA
result = minimize(cost, x0_scaled, method='COBYLA', constraints=constraints, options=options)

# Rescale the results
opt_scaled = result.x
opt = [opt_scaled[i] * scale_factors[i] for i in range(4)]

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

# Close Aspen Plus
Application.Quit()

# Convert x_values to a format that can be plotted (split QN1, QN2, QC)
QN1_values = [x[0] for x in x_values]
QN2_values = [x[1] for x in x_values]
QC_values = [x[2] for x in x_values]
SF_values = [x[3] for x in x_values]

# Create a single figure with 3 subplots
fig = plt.figure(figsize=(36, 12))

# Custom angles for the plots
elev_angle = 30 # Elevation angle (default is 30 degrees)
azim_angle = 130 # Azimuth angle (default is 120 degrees)

# Padding value between axis tick values and axis titles
axis_labelpad = 10
axis_titlepad = 10

# Plot 1: QN1 vs QN2 vs Objective Function
ax1 = fig.add_subplot(231, projection='3d')
ax1.plot(QN1_values, QN2_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax1.scatter(opt[0], opt[1], cost_min, color='blue', s=100, label='Optimal Point')
ax1.set_xlabel('QN1', labelpad=axis_labelpad)
ax1.set_ylabel('QN2', labelpad=axis_labelpad)
ax1.set_zlabel('Cost', labelpad=axis_labelpad)
ax1.set_title('QN1 vs QN2 vs Cost', pad=axis_titlepad)
ax1.view_init(elev=elev_angle, azim=azim_angle)
ax1.legend()

# Plot 2: QN1 vs QC vs Objective Function
ax2 = fig.add_subplot(232, projection='3d')
ax2.plot(QN1_values, QC_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax2.scatter(opt[0], opt[2], cost_min, color='blue', s=100, label='Optimal Point')
ax2.set_xlabel('QN1', labelpad=axis_labelpad)
ax2.set_ylabel('QC', labelpad=axis_labelpad)
ax2.set_zlabel('Cost', labelpad=axis_labelpad)
ax2.set_title('QN1 vs QC vs Cost', pad=axis_titlepad)
ax2.view_init(elev=elev_angle, azim=azim_angle)
ax2.legend()

# Plot 3: QN2 vs QC vs Objective Function
ax3 = fig.add_subplot(233, projection='3d')
ax3.plot(QN2_values, QC_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax3.scatter(opt[1], opt[2], cost_min, color='blue', s=100, label='Optimal Point')
ax3.set_xlabel('QN2', labelpad=axis_labelpad)
ax3.set_ylabel('QC', labelpad=axis_labelpad)
ax3.set_zlabel('Cost', labelpad=axis_labelpad)
ax3.set_title('QN2 vs QC vs Cost', pad=axis_titlepad)
ax3.view_init(elev=elev_angle, azim=azim_angle)
ax3.legend()

# Plot 4: QN1 vs SF vs Objective Function
ax4 = fig.add_subplot(234, projection='3d')
ax4.plot(QN1_values, SF_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax4.scatter(opt[0], opt[3], cost_min, color='blue', s=100, label='Optimal Point')
ax4.set_xlabel('QN1', labelpad=axis_labelpad)
ax4.set_ylabel('SF', labelpad=axis_labelpad)
ax4.set_zlabel('Cost', labelpad=axis_labelpad)
ax4.set_title('QN1 vs SF vs Cost', pad=axis_titlepad)
ax4.view_init(elev=elev_angle, azim=azim_angle)
ax4.legend()

# Plot 5: QN2 vs SF vs Objective Function
ax5 = fig.add_subplot(235, projection='3d')
ax5.plot(QN2_values, SF_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax5.scatter(opt[1], opt[3], cost_min, color='blue', s=100, label='Optimal Point')
ax5.set_xlabel('QN2', labelpad=axis_labelpad)
ax5.set_ylabel('SF', labelpad=axis_labelpad)
ax5.set_zlabel('Cost', labelpad=axis_labelpad)
ax5.set_title('QN2 vs SF vs Cost', pad=axis_titlepad)
ax5.view_init(elev=elev_angle, azim=azim_angle)
ax5.legend()

# Plot 6: QC vs SF vs Objective Function
ax6 = fig.add_subplot(236, projection='3d')
ax6.plot(QC_values, SF_values, objective_values, color='red', linestyle='-', marker='o', label='Optimization Path')
ax6.scatter(opt[2], opt[3], cost_min, color='blue', s=100, label='Optimal Point')
ax6.set_xlabel('QC', labelpad=axis_labelpad)
ax6.set_ylabel('SF', labelpad=axis_labelpad)
ax6.set_zlabel('Cost', labelpad=axis_labelpad)
ax6.set_title('QC vs SF vs Cost', pad=axis_titlepad)
ax6.view_init(elev=elev_angle, azim=azim_angle)
ax6.legend()

# Adjust the overall layout with margins to avoid trimming
plt.subplots_adjust(left=0.05, right=0.95, top=0.90, bottom=0.10, wspace=0.3)

# Save the combined figure with the script name
figure_name = script_name + '_3d_plots.png'
figure_path = os.path.join(os.getcwd(), figure_name)
plt.savefig(figure_path, bbox_inches='tight')  # Ensure everything fits in the saved image

# Show the figure (optional)
plt.show()

print(f'3D plots saved as: {figure_path}')
