import os
import win32com.client as win32

file = r"RECAP_revJ\RECAP_revJ.bkp"
aspen_Path = os.path.abspath(file)

print('Connecting to the Aspen Plus... Please wait ')
Application = win32.Dispatch('Apwn.Document') # Registered name of Aspen Plus
print('Connected!')

Application.InitFromArchive2(aspen_Path)
Application.visible = 0

# Objective function to minimize
def cost(x):
    N640QN = x
    return N640QN

# Constraint 1
def constraint1(x):
    cH2S = x - 0.00010
    return cH2S

# Initial guess
x0 = [550000]

# Constraints as a dictionary
constraints = [{'type': 'ineq', 'fun': constraint1}]

# Bounds
bounds = [(520000, 600000)]

# Solving the optimization problem
result = minimize(cost, x0, method='SLSQP', bounds=bounds, constraints=constraints)

# Output results
N640QN_opt = result.x
cost_min = result.fun
print(N640QN_opt, cost_min)