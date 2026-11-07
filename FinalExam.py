
# Name: Kerri Kuhnke
# Student ID: KerKuh7439
# Date: November 7, 2026
# Description: Final Exam Performance Assessment
# This program adds five numbers, collects income data,
# appends to a CSV file, and creates two charts.

import csv
import os
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl

from datetime import datetime
from openpyxl.chart import PieChart, Reference

# Change the working directory to FinalExam.
os.chdir("C:\\FinalExam")

# Store the student ID.
studentID = "KerKuh7439"

# Get today's date.
todayDate = datetime.now().strftime("%B %d, %Y")


# Function to ask the user for five numbers.
def askUser():

    # Initialize the total.
    total = 0

    # Loop five times to collect five numbers.
    # Each number is converted to an integer and
    # added to the running total.
    for count in range(5):

        number = int(input("Please enter a number: "))

        total = total + number

    # Display the sum of the five numbers.
    print("The sum for the 5 numbers entered is:", total)


# Function to collect five names and incomes.
def askIncome():

    # Open the existing CSV file in append mode.
    with open("final.csv", "a", newline="") as csvfile:

        # Create a CSV writer.
        csvWriter = csv.writer(csvfile)

        # Loop five times to collect five names and incomes.
        # Each entry is added to the existing CSV file
        # without deleting the original four entries.
        for count in range(5):

            name = input("Please enter a name: ")

            income = int(input("Please enter their income: "))

            # Write the name and income to the CSV file.
            csvWriter.writerow([name, income])

    print("\nFive new entries have been added to final.csv.")


# Function to create an Excel pie chart.
def excelPie():

    # Read the CSV file using Pandas.
    df = pd.read_csv("final.csv", header=None,
                     names=["Name", "Income"])

    # Convert income values to integers.
    df["Income"] = df["Income"].astype(int)

    # Create a new Excel workbook.
    wb = openpyxl.Workbook()

    # Select the active worksheet.
    ws = wb.active

    # Name the worksheet.
    ws.title = "Income Data"

    # Write the column headings.
    ws["A1"] = "Name"
    ws["B1"] = "Income"

    # Add the CSV data to the Excel worksheet.
    for row in df.itertuples(index=False, name=None):
        ws.append(row)

    # Create a PieChart object.
    myPieChart = PieChart()

    # Set the pie chart title to student ID and date.
    myPieChart.title = studentID + " " + todayDate

    # Select the numerical income values for the chart.
    myData = Reference(ws, min_col=2, min_row=2,
                       max_row=ws.max_row)

    # Select the names to label the pie chart.
    myDataLabels = Reference(ws, min_col=1, min_row=2,
                             max_row=ws.max_row)

    # Add the income data to the pie chart.
    myPieChart.add_data(myData)

    # Assign each person's name to the chart.
    myPieChart.set_categories(myDataLabels)

    # Set the height of the pie chart.
    myPieChart.height = 12

    # Set the width of the pie chart.
    myPieChart.width = 18

    # Insert the pie chart into the Excel worksheet.
    ws.add_chart(myPieChart, "A12")

    # Save the Excel workbook containing the chart.
    wb.save("final.xlsx")

    print("\nExcel pie chart saved to final.xlsx.")


# Function to create a vertical bar graph.
def verticalBar():

    # Read the CSV file using Pandas.
    df = pd.read_csv("final.csv", header=None,
                     names=["Name", "Income"])

    # Convert income values to integers.
    df["Income"] = df["Income"].astype(int)

    # Create a vertical bar graph.
    df.plot(x="Name", y="Income",
            kind="bar", color="green")

    # Set the graph title.
    plt.title(studentID + " " + todayDate)

    # Label the horizontal axis.
    plt.xlabel("Name")

    # Label the vertical axis.
    plt.ylabel("Income")

    # Keep the names horizontal.
    plt.xticks(rotation=0)

    # Adjust the graph layout.
    plt.tight_layout()

    # Display the vertical bar graph.
    plt.show()


# Main function to run the program.
def main():

    # Display student ID.
    print(studentID)

    # Display the current system date and time.
    print("Date and Time:", datetime.now())

    # Ask the user to enter five numbers.
    askUser()

    # Ask for five names and incomes.
    askIncome()

    # Create the Excel pie chart.
    excelPie()

    # Create the vertical bar graph.
    verticalBar()


# Start the program.
main()
