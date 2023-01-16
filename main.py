import calendar
import os
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import create_graphs as cg

TOTAL_MONTHLY_HOURS = 0
EMPLOYEES_PER_GRAPH = 4


def get_month_number(month_name):
    months = {
        month.lower(): index for index, month in enumerate(calendar.month_name) if month
    }
    return months[month_name.lower()]


def main():
    # Read the excel file
    allocation_sheet_df = pd.read_excel(
        os.path.join(os.getcwd(), "./Main.xlsx"), sheet_name="Allocation"
    )
    employee_sheet_df = pd.read_excel(
        os.path.join(os.getcwd(), "./Main.xlsx"), sheet_name="Employees"
    )
    project_sheet_df = pd.read_excel(
        os.path.join(os.getcwd(), "./Main.xlsx"), sheet_name="Projects"
    )
    variables_df = pd.read_excel(
        os.path.join(os.getcwd(), "./Main.xlsx"),
        sheet_name="Variables",
        header=0,
    )

    # Reading variables and setting constants.
    EMPLOYEES_PER_GRAPH = variables_df.iloc[0]["Value"]
    TOTAL_MONTHLY_HRS = variables_df.iloc[1]["Value"]

    monthly_hours = {calendar.month_name[i]: TOTAL_MONTHLY_HRS for i in range(1, 13)}
    monthly_hours = pd.Series(monthly_hours)

    # filtering Data
    allocation_sheet_df["Month"] = [
        get_month_number(i) for i in allocation_sheet_df["Month"]
    ]

    allocation_sheet_df.dropna(inplace=True)
    allocation_sheet_df.drop_duplicates(inplace=True)

    # Creating Employee Dictionary to then create the dataframes
    employees = {}  # this holds dictionaries for each employee.

    for i in range(employee_sheet_df.__len__()):
        employee = {
            "name": employee_sheet_df.iloc[i].values[0],
            "branch": employee_sheet_df.iloc[i].values[1],
            "projects": {},
        }
        employees[employee_sheet_df.iloc[i].values[0]] = employee

    # Adding projects to each employee
    for _, i in enumerate(employees):
        emp_df = allocation_sheet_df[i == allocation_sheet_df["Name of Employee"]]
        # adding projects
        for j in range(len(emp_df)):
            employees[i]["projects"][emp_df.iloc[j]["Project Name"]] = {}
            employees[i]["projects"][emp_df.iloc[j]["Project Name"]] = {
                k: v
                for (k, v) in zip(
                    [calendar.month_name[i] for i in range(1, 13)],
                    [0 for i in range(1, 13)],
                )
            }
        # assigning time for each project
        for j in range(len(emp_df)):
            employees[i]["projects"][emp_df.iloc[j]["Project Name"]][
                calendar.month_name[emp_df["Month"].iloc[j]]
            ] += emp_df.iloc[j]["Number of Hours"]

        # Finally creating the Employee Dataframe, this is the main thing and has all the information that we need.
        employee_df = pd.DataFrame(employees).transpose()

    # Creating Graphs

    these_months = [
        calendar.month_name[i]
        for i in range(datetime.now().month, datetime.now().month + 4)
    ]

    plt.style.use("default")

    # ---------------------------------------- gotta put for loop here ------------------------
    start = 0
    while start <= employee_df.__len__():
        current_employees_df = employee_df.iloc[start : start + EMPLOYEES_PER_GRAPH]
        print(current_employees_df)
        print(EMPLOYEES_PER_GRAPH)
        start += EMPLOYEES_PER_GRAPH
        projects_list = []

        # get the projects that the current employees are working on.
        for i in current_employees_df.iterrows():
            projects_list.extend(i[1].projects.keys())

        projects_list = list(set(projects_list))
        projects_list.append("free_time")

        category_colors = cg.createColors(projects_list)

        # now well create dataframes for each person for each project.
        # these dataframes have the hours for each month for each project.

        employee_project_dataframes = {}
        employee_project_dataframes_cumsum = {}

        for i in current_employees_df.iterrows():
            current_employee_df = pd.DataFrame(
                index=projects_list, columns=these_months
            )
            for j in i[1].projects.keys():
                current_employee_df.loc[j] = i[1].projects[j]
            current_employee_df.fillna(0, inplace=True)
            current_employee_df.loc[
                "free_time"
            ] = monthly_hours - current_employee_df.sum(axis=0)
            for j in current_employee_df.columns:
                current_employee_df[j] = np.array(current_employee_df[j])

            employee_project_dataframes[i[0]] = current_employee_df
            employee_project_dataframes[i[0]] = employee_project_dataframes[i[0]][
                employee_project_dataframes[i[0]].columns[::-1]
            ]
            employee_project_dataframes_cumsum[i[0]] = current_employee_df.cumsum(
                axis=0
            )
            employee_project_dataframes_cumsum[
                i[0]
            ] = employee_project_dataframes_cumsum[i[0]][
                employee_project_dataframes_cumsum[i[0]].columns[::-1]
            ]

        # making graphs now.

        labels = list(reversed(these_months))
        cg.makeGraph(
            labels,
            employee_project_dataframes,
            employee_project_dataframes_cumsum,
            projects_list,
            category_colors,
            start,
            EMPLOYEES_PER_GRAPH
        )
if(not os.path.exists(os.path.join(os.getcwd(), "Employee Graphs"))):
    os.mkdir(os.path.join(os.getcwd(), "Employee Graphs"))

main()
