# ===============================
# IMPORT LIBRARIES
# ===============================
from pulp import *
import pandas as pd

# ===============================
# DATA FROM PROBLEM
# ===============================
stations = 15
ground = [51,67,61,55,40,45,20,22,20,42,60,42,25,28,22]

width = 20
distance = 100
volume_factor = width * distance   # = 2000

cut_cost = 150
fill_cost = 100
dump_cost = 70
borrow_cost = 120

grade_limit = 5
grade_change_limit = 6

initial_grade = 4
final_grade = -3


# ===============================
# SOLVER FUNCTION
# ===============================
def solve_model(relaxed_segment=None):

    prob = LpProblem("Pipeline_Optimization", LpMinimize)

    # Variables
    y = LpVariable.dicts("Elevation", range(stations))
    cut = LpVariable.dicts("Cut", range(stations), lowBound=0)
    fill = LpVariable.dicts("Fill", range(stations), lowBound=0)

    borrow = LpVariable("Borrow", lowBound=0)
    dump = LpVariable("Dump", lowBound=0)

    g = LpVariable.dicts("Grade", range(stations-1))

    # Objective
    prob += volume_factor * (
        lpSum(cut_cost*cut[i] + fill_cost*fill[i] for i in range(stations))
        + dump_cost*dump + borrow_cost*borrow
    )

    # Cut-fill relation
    for i in range(stations):
        prob += y[i] - ground[i] == fill[i] - cut[i]

    # Material balance
    prob += lpSum(cut[i] for i in range(stations)) + borrow == \
            lpSum(fill[i] for i in range(stations)) + dump

    # Fixed endpoints
    prob += y[0] == ground[0]
    prob += y[stations-1] == ground[stations-1]

    # Grade definition
    for i in range(stations-1):
        prob += g[i] == y[i+1] - y[i]

    # Grade constraints
    for i in range(stations-1):

        limit = grade_limit
        if relaxed_segment == i:
            limit = 5.5   # 0.5% increase

        prob += g[i] <= limit
        prob += g[i] >= -limit

    # Grade change constraints
    prob += g[0] - initial_grade <= grade_change_limit
    prob += initial_grade - g[0] <= grade_change_limit

    for i in range(1, stations-1):
        prob += g[i] - g[i-1] <= grade_change_limit
        prob += g[i-1] - g[i] <= grade_change_limit

    prob += final_grade - g[stations-2] <= grade_change_limit
    prob += g[stations-2] - final_grade <= grade_change_limit

    # Solve
    prob.solve(PULP_CBC_CMD(msg=0))

    return value(prob.objective)


# ===============================
# BASE SOLUTION
# ===============================
base_cost = solve_model()

print(f"Base Minimum Cost = ₹{base_cost:,.2f}")


# ===============================
# SENSITIVITY ANALYSIS
# ===============================
results = []

for seg in range(stations-1):

    new_cost = solve_model(relaxed_segment=seg)
    saving = base_cost - new_cost

    results.append({
        "Segment": f"{seg+1}-{seg+2}",
        "Base Cost": base_cost,
        "New Cost (5.5%)": new_cost,
        "Cost Saving": saving
    })


# ===============================
# CREATE DATAFRAME
# ===============================
df = pd.DataFrame(results)

# Best segment
best_row = df.loc[df["Cost Saving"].idxmax()]


# ===============================
# EXPORT TO EXCEL
# ===============================
file_name = "pipeline_sensitivity_analysis.xlsx"

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:

    # Sheet 1: Full analysis
    df.to_excel(writer, sheet_name="Sensitivity Analysis", index=False)

    # Sheet 2: Best segment
    pd.DataFrame([best_row]).to_excel(writer, sheet_name="Best Segment", index=False)


print("\nExcel file generated:", file_name)

print("\nBest Segment:")
print(best_row)