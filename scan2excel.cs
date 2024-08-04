using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;

public class MainForm : Form
{
    private Label myLabel;
    private TextBox myTextBox;
    private Button buttonExport;
    private Button buttonReportDaily;
    private Button buttonReportWeekly;
    private Button buttonReportMonthly;
    private Button buttonInfo;
    private Button buttonSettings;
    private Button buttonLicense;

    public MainForm()
    {
        // Initialize controls
        myLabel = new Label { Text = "Scan", Location = new System.Drawing.Point(10, 10) };
        myTextBox = new TextBox { Location = new System.Drawing.Point(10, 40), Width = 200 };
        buttonExport = new Button { Text = "Export", Location = new System.Drawing.Point(10, 70), Width = 100 };

        // Initialize new buttons
        buttonReportDaily = new Button { Text = "Report Daily", Size = new System.Drawing.Size(100, 30) };
        buttonReportWeekly = new Button { Text = "Report Weekly", Size = new System.Drawing.Size(100, 30) };
        buttonReportMonthly = new Button { Text = "Report Monthly", Size = new System.Drawing.Size(100, 30) };
        buttonInfo = new Button { Text = "Info", Size = new System.Drawing.Size(100, 30) };
        buttonSettings = new Button { Text = "Settings", Size = new System.Drawing.Size(100, 30) };
        buttonLicense = new Button { Text = "License", Size = new System.Drawing.Size(100, 30) };

        // Set initial locations for the new buttons
        buttonReportDaily.Location = new System.Drawing.Point(this.ClientSize.Width - buttonReportDaily.Width - 10, this.ClientSize.Height - buttonReportDaily.Height - 90);
        buttonReportWeekly.Location = new System.Drawing.Point(this.ClientSize.Width - buttonReportWeekly.Width - 10, this.ClientSize.Height - buttonReportWeekly.Height - 50);
        buttonReportMonthly.Location = new System.Drawing.Point(this.ClientSize.Width - buttonReportMonthly.Width - 10, this.ClientSize.Height - buttonReportMonthly.Height - 10);
        buttonLicense.Location = new System.Drawing.Point(10, this.ClientSize.Height - buttonLicense.Height - 90);
        buttonInfo.Location = new System.Drawing.Point(10, this.ClientSize.Height - buttonInfo.Height - 50);
        buttonSettings.Location = new System.Drawing.Point(10, this.ClientSize.Height - buttonSettings.Height - 10);

        // Anchor buttons to the bottom right
        buttonReportDaily.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        buttonReportWeekly.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        buttonReportMonthly.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        buttonLicense.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        buttonInfo.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        buttonSettings.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

        // Add event handlers
        buttonExport.Click += ButtonExport_Click;
        buttonLicense.Click += ButtonLicense_Click;
        buttonInfo.Click += ButtonInfo_Click;

        // Add controls to the form
        Controls.Add(myLabel);
        Controls.Add(myTextBox);
        Controls.Add(buttonExport);
        Controls.Add(buttonReportDaily);
        Controls.Add(buttonReportWeekly);
        Controls.Add(buttonReportMonthly);
        Controls.Add(buttonLicense);
        Controls.Add(buttonInfo);
        Controls.Add(buttonSettings);

        // Set form properties
        this.Size = new System.Drawing.Size(400, 450); // Ensure the form is large enough to display all controls
    }

    private void ButtonExport_Click(object? sender, EventArgs e)
    {
        string data = myTextBox.Text; // Get the text from the textbox.
        if (IsValidInput(data))
        {
            SaveDataToExcel(data); // Call the method to save the data to Excel.
        }
        else
        {
            MessageBox.Show("Invalid Input. No special characters allowed."); // Show the message box if input is invalid.
        }
    }

    private bool IsValidInput(string input)
    {
        // Regular expression to match only alphanumeric characters and spaces.
        Regex regex = new Regex("^[a-zA-Z0-9 ]*$");
        return regex.IsMatch(input);
    }

    private void SaveDataToExcel(string data)
    {
        string filePath = @"E:\DVISTLtest.xlsx"; // Define the file path for the Excel file.
        FileInfo fileInfo = new FileInfo(filePath); // Create a FileInfo object for the file.

        using (ExcelPackage package = new ExcelPackage(fileInfo)) // Open the Excel package.
        {
            // Handle the "REPO" worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets["REPO"]; // Get the worksheet named "REPO".
            if (worksheet == null) // If the worksheet does not exist,
            {
                worksheet = package.Workbook.Worksheets.Add("REPO"); // Create a new worksheet named "REPO".
                worksheet.Cells[1, 1].Value = "Data"; // Set the header for column A.
                worksheet.Cells[1, 2].Value = "Count"; // Set the header for column B.
            }

            bool matchFound = false; // Initialize a flag to check if a match is found.
            int rowCount = worksheet.Dimension?.Rows ?? 0; // Get the number of rows in the worksheet.

            for (int row = 2; row <= rowCount; row++) // Loop through each row starting from the second row.
            {
                if (worksheet.Cells[row, 1].Text == data) // If the data in column A matches the input data,
                {
                    double currentValue = worksheet.Cells[row, 2].GetValue<double>(); // Get the current value in column B.
                    worksheet.Cells[row, 2].Value = (int)currentValue + 1; // Increment the value in column B.
                    matchFound = true; // Set the flag to indicate a match is found.
                    break; // Exit the loop.
                }
            }

            if (!matchFound) // If no match is found,
            {
                int newRow = rowCount + 1; // Determine the next available row.
                worksheet.Cells[newRow, 1].Value = data; // Set the input data in column A.
                worksheet.Cells[newRow, 2].Value = 1; // Set the count to 1 in column B.
            }

            // Handle the "LOG" worksheet
            ExcelWorksheet logWorksheet = package.Workbook.Worksheets["LOG"]; // Get the worksheet named "LOG".
            if (logWorksheet == null) // If the worksheet does not exist,
            {
                logWorksheet = package.Workbook.Worksheets.Add("LOG"); // Create a new worksheet named "LOG".
                logWorksheet.Cells[1, 1].Value = "Data"; // Set the header for column A.
                logWorksheet.Cells[1, 2].Value = "User"; // Set the header for column B.
                logWorksheet.Cells[1, 3].Value = "Timestamp"; // Set the header for column C.
            }

            int logRowCount = logWorksheet.Dimension?.Rows ?? 0; // Get the number of rows in the "LOG" worksheet.
            int newLogRow = logRowCount + 1; // Determine the next available row in the "LOG" worksheet.
            logWorksheet.Cells[newLogRow, 1].Value = data; // Set the input data in column A.
            logWorksheet.Cells[newLogRow, 2].Value = Environment.UserName; // Set the logged-in user in column B.
            logWorksheet.Cells[newLogRow, 3].Value = DateTime.Now.ToString(); // Set the current system time in column C.

            // Save the package to the file
            package.Save(); // Save the changes to the Excel file.
            MessageBox.Show("Data exported to " + filePath); // Show a message box indicating the data has been exported.
        }
    }

    private void ButtonLicense_Click(object? sender, EventArgs e)
    {
        // Create a new form to display the license
        Form licenseForm = new Form
        {
            Text = "License",
            Size = new System.Drawing.Size(400, 300)
        };

        // Create a label to display the license text
        Label licenseLabel = new Label
        {
            Text = "GNU GENERAL PUBLIC LICENSE\n Version 3, 29 June 2007\n Walter Migaud\n https://github.com/waltermigaud/scan2excel/ \n This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.\n\nThis program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.\n\nYou should have received a copy of the GNU General Public License along with this program. If not, see http://www.gnu.org/licenses/.",
            Dock = DockStyle.Fill,
            TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        };

        // Add the label to the new form
        licenseForm.Controls.Add(licenseLabel);

        // Show the new form
        licenseForm.ShowDialog();
    }

    private void ButtonInfo_Click(object? sender, EventArgs e)
    {
        // Create a new form to display the info
        Form infoForm = new Form
        {
            Text = "Info",
            Size = new System.Drawing.Size(400, 300)
        };

        // Create a label to display the info text
        Label infoLabel = new Label
        {
            Text = "This program aims to provide a feature for connecting tag scanning to a local or networked excel document with a modest amount of customizable functionality and reporting options. Different iterations addressing concerns of functionality and security are in development. For custom solutions and features please contact me at my github profilename, at outlook. August 4, 2024",
            Dock = DockStyle.Fill,
            TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        };

        // Add the label to the new form
        infoForm.Controls.Add(infoLabel);

        // Show the new form
        infoForm.ShowDialog();
    }

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}