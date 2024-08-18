import pandas as pd
import numpy as np
import openpyxl

# Load the data from the Excel file
data = pd.read_excel("school_data.xlsx")


# Function to normalize a series
def normalize(series):
    return series / np.sqrt(np.sum(series**2))


# Normalize relevant columns
data["Norm_TS"] = normalize(data["Test Scores"])
data["Norm_GR"] = normalize(data["Graduation Rate"])
data["Norm_TSR"] = normalize(data["Teacher-Student Ratio"])
data["Norm_LM"] = normalize(data["Learning Materials"])
data["Norm_FP"] = normalize(data["Facility Provided"])
data["Norm_ELL"] = normalize(data["English Language Learners"])

# Printing Normalised data
print("\n", data)

# Define weights
weights = {
    "Norm_TS": 0.40,
    "Norm_GR": 0.40,
    "Norm_TSR": 0.30,
    "Norm_LM": 0.30,
    "Norm_FP": 0.30,
    "Norm_ELL": 0.30,
}
# calculation of value
for key in weights:
    data[key] = data[key] * weights[key]
    print("\n",data[key])

# Combine weighted columns into a numpy array for easier manipulation
criteria_matrix = data[
    ["Norm_TS", "Norm_GR", "Norm_TSR", "Norm_LM", "Norm_FP", "Norm_ELL"]
].values.T

# Determine Ideal (A*) and Negative-Ideal (A-) Solutions
ideal_solution = np.max(criteria_matrix, axis=1)
negative_ideal_solution = np.min(criteria_matrix, axis=1)
print("\nideal", ideal_solution)
print("\nnegative-ideal", negative_ideal_solution)


# Calculate the Euclidean distance from ideal and negative-ideal solutions
def euclidean_distance(matrix, solution):
    return np.sqrt(np.sum((matrix - solution.reshape(-1, 1)) ** 2, axis=0))

# Calculate the measures
dist_from_ideal = euclidean_distance(criteria_matrix, ideal_solution)
dist_from_neg_ideal = euclidean_distance(criteria_matrix, negative_ideal_solution)
print("\nDistance from ideal", dist_from_ideal)
print("\nDistance from negative-ideal", dist_from_neg_ideal)


# Calculate the Relative Closeness to the Ideal Solution
relative_closeness = dist_from_neg_ideal / (dist_from_ideal + dist_from_neg_ideal)
print("\nRelative closeness", relative_closeness)

# Rank schools based on relative closeness
data["Relative_Closeness"] = relative_closeness
data["Rank"] = data["Relative_Closeness"].rank(ascending=True)
print(data)
data = data.sort_values(by="Rank")
print("\nSorted data by Ranking\n",data)
    
    
def find_closest(numbers, target):
    closest_value = min(numbers, key=lambda x: abs(x - target))
    return closest_value
    
    
def getMostRequired(col_name, col_num):
    cval=find_closest(data[col_name],negative_ideal_solution[col_num])
    filtered_rows = data[data[col_name] == cval]
    return filtered_rows

def getLeastRequired(col_name, col_num):
    cval=find_closest(data[col_name],ideal_solution[col_num])
    filtered_rows = data[data[col_name] == cval]
    return filtered_rows


# Getting School which needs staff the most
print("\nSchool which needs the staff most is: \n")
filt_row = getMostRequired("Norm_TSR", 2)
print(filt_row)


# Getting School which needs staff the least
print("\nSchool which needs the staff least is: \n")
filt_row = getLeastRequired("Norm_TSR", 2)
print(filt_row)


# Getting School which needs materials the most
print("\nSchool which needs the material most is: \n")
filt_row = getMostRequired("Norm_LM", 3)
print(filt_row)

# Getting School which needs Materials the least
print("\nSchool which needs the Materials the least is: \n")
filt_row = getLeastRequired("Norm_LM", 3)
print(filt_row)

# Getting School which needs funding the most
print("\nSchool which needs the funding most is: \n")
filt_row = getMostRequired("Norm_FP", 4)
print(filt_row)


# Getting School which needs funding the least
print("\nSchool which needs the funding least is: \n")
filt_row = getLeastRequired("Norm_FP", 4)
print(filt_row)



# Save the results to a new Excel file
with pd.ExcelWriter("school_rankings_topsis.xlsx", engine="openpyxl") as writer:
     data.to_excel(writer, sheet_name="Rankings", index=False)

# Display the ranked data