using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

public class MainForm : Form
{
    private Label myLabel;
    private TextBox myTextBox;
    private Button buttonExport;

    public MainForm()
    {
        myLabel = new Label { Text = "Scan Status", Location = new System.Drawing.Point(10, 10) };
        myTextBox = new TextBox { Location = new System.Drawing.Point(10, 40) };
        buttonExport = new Button { Text = "Export", Location = new System.Drawing.Point(10, 70) };

        buttonExport.Click += ButtonExport_Click;

        Controls.Add(myLabel);
        Controls.Add(myTextBox);
        Controls.Add(buttonExport);
    }

    private void ButtonExport_Click(object? sender, EventArgs e)
    {
        string data = myTextBox.Text; // Get the text from the textbox.
        SaveDataToExcel(data); // Call the method to save the data to Excel.
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

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}