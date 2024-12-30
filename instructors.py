!pip install pulp
import pandas as pd
import pulp
# Load the data from provided Excel sheets
courses_path = 'Courses.xlsx'
instructors_path = 'Instructors.xlsx'
preferences_path = 'Preferences.xlsx'

courses_df = pd.read_excel(courses_path)
instructors_df = pd.read_excel(instructors_path)
preferences_df = pd.read_excel(preferences_path)

# Extract data
courses = courses_df['course_name'].tolist()
instructors = instructors_df['instructor_name'].tolist()
qualification_columns = courses_df.columns[2:].tolist()  # Get qualifications from the header row (A2-A4)

# Load the workload unit data
courses_df = pd.read_excel(courses_path)  # Reload to include the new column
course_workload_units = {course: courses_df.loc[courses_df['course_name'] == course, 'Workload Unit'].values[0]
                         for course in courses}
# Create dictionaries for qualifications, seniority, workload, and preferences
qualifications = {}
for instructor in instructors:
    for course in courses:
        # Check if the instructor meets the qualification for the course based on subject requirements
        course_requirements = courses_df.loc[courses_df['course_name'] == course, qualification_columns].values[0]
        instructor_qualifications = instructors_df.loc[instructors_df['instructor_name'] == instructor, qualification_columns].values[0]
        # Check if the instructor meets all required qualifications for the course
        is_qualified = all(instructor_qualifications[i] >= course_requirements[i] for i in range(len(qualification_columns)))
        qualifications[(instructor, course)] = 1 if is_qualified else 0

seniority = {instructor: instructors_df.loc[instructors_df['instructor_name'] == instructor, 'Seniority'].values[0] for instructor in instructors}


workload = {instructor: instructors_df.loc[instructors_df['instructor_name'] == instructor, 'Workload'].values[0] for instructor in instructors}

preferences = {(instructor, course): preferences_df.loc[preferences_df['instructor_name'] == instructor, course].values[0]
                for instructor in instructors for course in courses}

print(preferences_df.columns) # Print the column names
print(preferences_df.head())  # Print the first few rows of data

# Define the optimization problem
model = pulp.LpProblem("Instructor_Course_Assignment", pulp.LpMaximize)

# Define decision variables
x = pulp.LpVariable.dicts("assignment", (instructors, courses), cat='Binary')

# Revised Objective: Prioritize preferences more heavily by separating it from seniority and qualifications
model += pulp.lpSum(2 * seniority[i] * qualifications[i, j] * x[i][j] + 2 * preferences[i, j] * qualifications[i, j] * x[i][j]
                    for i in instructors for j in courses)

# Constraint 1: Instructor assigned workload <= workload
for i in instructors:
    model += pulp.lpSum(course_workload_units[j] * x[i][j] for j in courses) <= workload[i], f"Workload_{i}"

# Constraint 2: No more than one instructor per course
for j in courses:
    model += pulp.lpSum(x[i][j] for i in instructors) <= 1, f"One_Instructor_Per_Course_{j}"

# Solve the problem
model.solve(pulp.PULP_CBC_CMD(timeLimit=60))

# Prepare and save the results to Excel
assignments = []
for i in instructors:
    for j in courses:
        if pulp.value(x[i][j]) == 1:
            assignments.append({'Instructor': i, 'Course': j})

assignments_df = pd.DataFrame(assignments)
assignments_df.to_excel('course_assignments.xlsx', index=False)

# Display the results
for i in instructors:
    for j in courses:
        if pulp.value(x[i][j]) == 1:
            print(f"Instructor {i} is assigned to course {j}")

