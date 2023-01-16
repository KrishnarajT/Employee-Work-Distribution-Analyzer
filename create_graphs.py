import calendar
import os
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

plt.style.use("default")


def createColors(projects_list):

    category_colors_1 = plt.colormaps["RdYlGn"](
        np.linspace(0.15, 1, projects_list.__len__() - 1)
    )

    # for free time
    category_color_green = [
        [0.429527104959631, 0.7540945790080739, 0.3914733082901554, 1.0]
    ]

    colors_series = pd.Series(category_colors_1.tolist())

    # Shuffle the colors
    colors_series = colors_series.sample(frac=1)

    # concat the green color to the end of the series
    colors_series = pd.concat([colors_series, pd.Series(category_color_green)])
    category_colors = pd.Series([i for i in colors_series], index=projects_list)

    return category_colors


def makeGraph(
    labels,
    employee_project_dataframes,
    employee_project_dataframes_cumsum,
    projects_list,
    category_colors,
    batch,
    EMPLOYEE_PER_GRAPH=4,
):
    first_one_done = False
    # making 4 subplots
    fig, axes = plt.subplots(
        employee_project_dataframes.__len__(), sharex=True, sharey=True, figsize=(15, 5)
    )
    plt.xlabel("Number of Hours")
    plt.subplots_adjust(hspace=0.1)

    # iterating through 4 subplots, each subplot is a different employee
    for ax, emp_names in zip(axes, employee_project_dataframes_cumsum.keys()):

        # remove ticks on both axes for employees.
        ax.tick_params(
            axis="x", which="both", bottom=False, top=False, labelbottom=False
        )

        print(emp_names, "---------------------------------")
        ax.set_ylabel(emp_names, rotation=0, labelpad=50)

        # iterating through each project for that employee, there will be a few projects, and the y axis would have the months
        for i, (colname, color) in enumerate(zip(projects_list, category_colors)):
            widths = employee_project_dataframes[emp_names].loc[colname]
            starts = employee_project_dataframes_cumsum[emp_names].loc[colname] - widths
            rects = ax.barh(labels, widths, left=starts, label=colname, color=color)
            r, g, b, _ = color
            text_color = "white" if r * g * b < 0.2 else "black"
            # print rects
            rects.datavalues = [np.nan if i == 0 else i for i in rects.datavalues]
            # print(rects.datavalues)
            ax.bar_label(
                rects,
                labels=rects.datavalues,
                label_type="center",
                color=text_color,
                padding=1,
                fmt="%.1f",
            )
        if not first_one_done:
            ax.legend(
                ncol=len(projects_list),
                loc="upper center",
                fontsize="large",
                bbox_to_anchor=(0.5, 1.6),
            )
            first_one_done = True

    # save the graph
    fig.savefig(
        os.path.join(
            os.getcwd(),
            "Employee Graphs",
            f"Employee Graph - {batch/EMPLOYEE_PER_GRAPH}.png",
        ),
        bbox_inches="tight",
        format="png",
        dpi=300,
    )
    plt.close(fig)
    print("saved figure! ", batch)
