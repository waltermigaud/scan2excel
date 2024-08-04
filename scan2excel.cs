using System;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

public class MainForm : Form
{
    private TextBox myTextBox = new TextBox();
    private Label myLabel = new Label();
    private System.Windows.Forms.Timer delayTimer = new System.Windows.Forms.Timer();
    private Button buttonExport = new Button();

    public MainForm()
    {
        // Set properties for the TextBox
        myTextBox.Location = new System.Drawing.Point(100, 50);
        myTextBox.TextChanged += new EventHandler(MyTextBox_TextChanged);

        // Set properties for the label
        myLabel.Text = "Scan DataMatrix on tool base.";
        myLabel.Location = new System.Drawing.Point(50, 100);
        myLabel.AutoSize = true;

        // Set properties for the Timer
        delayTimer.Interval = 2000; // 2 seconds
        delayTimer.Tick += new EventHandler(DelayTimer_Tick);

        // Set properties for the Button
        buttonExport.Text = "Export";
        buttonExport.Location = new System.Drawing.Point(100, 150);
        buttonExport.Click += new EventHandler(ButtonExport_Click);

        // Add components to the form
        Controls.Add(myTextBox);
        Controls.Add(myLabel);
        Controls.Add(buttonExport);
    }

    private void MyTextBox_TextChanged(object? sender, EventArgs e)
    {
        delayTimer.Stop(); // Stop any previous timer
        delayTimer.Start(); // Start a new timer
    }

    private void DelayTimer_Tick(object? sender, EventArgs e)
    {
        delayTimer.Stop();
        if (!string.IsNullOrEmpty(myTextBox.Text)) // detects ANY input which waits 2 seconds per timer properties above
        {
            myLabel.Text = "Hell yea Scan Success";
        }
    }

    private void ButtonExport_Click(object? sender, EventArgs e)
    {
        string data = myTextBox.Text;
        SaveDataToExcel(data);
    }

    private void SaveDataToExcel(string data)
    {
        string filePath = @"E:\DVISTLtest.xlsx";
        FileInfo fileInfo = new FileInfo(filePath);

        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["REPO"];
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add("REPO");
                worksheet.Cells[1, 1].Value = "Data";
                worksheet.Cells[1, 2].Value = "Count";
            }

            bool matchFound = false;
            int rowCount = worksheet.Dimension?.Rows ?? 0;

            for (int row = 2; row <= rowCount; row++)
            {
                if (worksheet.Cells[row, 1].Text == data)
                {
                    double currentValue = worksheet.Cells[row, 2].GetValue<double>();
                    worksheet.Cells[row, 2].Value = (int)currentValue + 1;
                    matchFound = true;
                    break;
                }
            }

            if (!matchFound)
            {
                int newRow = rowCount + 1;
                worksheet.Cells[newRow, 1].Value = data;
                worksheet.Cells[newRow, 2].Value = 1;
            }

            // Save the package to the file
            package.Save();
            MessageBox.Show("Data exported to " + filePath);
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