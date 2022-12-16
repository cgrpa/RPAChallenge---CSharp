using System;
using System.Runtime.InteropServices;
using System.Threading;
using Accessibility;
using FlaUI;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using NUnit.Framework;
using System.Data;
using System.Data.OleDb;
using System.Data.Common;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using FlaUI.UIA3.Patterns;
using System.Collections.Generic;
using System.Diagnostics;

namespace RPAChallenge___CSharp
{
    class Program
    {
        public const int BigWaitTimeout = 3000;
        public const int SmallWaitTimeout = 1000;
        static void Main(string[] args)
        {
            ParallelOptions parallelOptions = new ParallelOptions();
            parallelOptions.MaxDegreeOfParallelism = Environment.ProcessorCount;
            // use max degree of parallelism.

            KillProcesses("msedge");
            string rpaChallengePath = Path.Combine(Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName, "challenge.xlsx");
            // Read DataTable
            DataTable data = ReadExcelFile("Sheet1", rpaChallengePath);

            //Launch application, force access accessibility
            var app = FlaUI.Core.Application.Launch(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", "--force-renderer-accessibility https://rpachallenge.com");
            //Replace your edge file path here if different.

            //Create new automation ensuring correct use of IDisposable objects
            using (var automation = new UIA3Automation())
            {
                var mainWindow = app.GetMainWindow(automation);
                //Define mainWindow

                FlaUI.Core.Input.Wait.UntilResponsive(mainWindow);
                //Wait until mainWindow is responsive

                mainWindow.Patterns.Window.Pattern.SetWindowVisualState(WindowVisualState.Maximized);
                //Maximise mainWindow


                // Determine control elements that do not change by Name attribute
                Button submitBtn = WaitForElement(() => mainWindow.FindFirstDescendant(cf => cf.ByName("Submit")).AsButton());
                Button startBtn = WaitForElement(() => mainWindow.FindFirstDescendant(cf => cf.ByName("START")).AsButton());

                var inputGroup = WaitForElement(() => mainWindow.FindFirstByXPath("/Pane[2]/Pane/Pane[1]/Pane[2]/Pane[2]/Pane/Pane/Pane/Pane[1]/Document/Group/Group[2]"));
                // This is the grouping of all input elements and their labels


                FlaUI.Core.Input.Wait.UntilResponsive(submitBtn);
                //Wait until Submit Button is responsive

                startBtn.Click(false); // Disable moving mouse


                //IQueryable vs IEnumerable: is IQueryable always better and faster? https://stackoverflow.com/q/43419228
                foreach (DataRow row in data.AsEnumerable().AsQueryable())
                {
                    AutomationElement[] inputGroupChildren = inputGroup.FindAllChildren(); //create Array of inputGroup Child elements

                    //Find and enter input data in parallel
                    Parallel.ForEach(data.Columns.Cast<DataColumn>().AsQueryable(), parallelOptions, (column, state, index) =>
                    {
                        int childIndex = -1;
                        string lblName = column.ColumnName.Trim();
                        string inputText = row[column.ColumnName].ToString();
                        childIndex = FindInputIndex(inputGroupChildren.AsQueryable<AutomationElement>(), lblName);

                        if (childIndex == -1) throw new SystemException("Could not find element for: " + lblName);

                        var inputElement = inputGroupChildren[childIndex + 1].AsTextBox(); // the Input element always comes after the label.

                        if (row[column.ColumnName].ToString() == string.Empty) throw new SystemException("Empty row in data"); // Just incase you have empty columns in your data.

                        inputElement.Text = inputText; // Set Text.
                    }
                    );

                    // Submit that row's data
                    submitBtn.Click(false);
                }
            }




        }

        // Kill all given processes.
        public static void KillProcesses(string processName) //ommit the .exe
        {
            foreach (var process in Process.GetProcessesByName(processName))
            {
                process.Kill();
            }
        }

        // Find index of element which contains lblText
        public static int FindInputIndex(IQueryable<AutomationElement> group, string lblText)
        {
            int i = 0;
            foreach (AutomationElement child in group)
            {
                if (child.FindAllChildren().Length > 0) // If the child has further children it may be an Input box.
                {
                    var lblChild = child.FindAllChildren()[0]; // The first element of child would be an 
                    if (lblChild.Name == lblText)
                    {
                        return i;
                    }
                }
                i++;

            }

            return -1;
        }



        //Stolen from: https://stackoverflow.com/a/58780421
        //This function uses the Retry class in FlaUI.Core.Tools.Retry to wait for an element to be ready.
        public static T WaitForElement<T>(Func<T> getter)
        {
            var retry = Retry.WhileNull<T>(
                    () => getter(),
                    TimeSpan.FromMilliseconds(BigWaitTimeout));

            if (!retry.Success)
            {
                Assert.Fail("Failed to get an element within a wait timeout");
            }
            return retry.Result;

        }



        // Read Excel File via OleDB
        static DataTable ReadExcelFile(string sheetName, string path)
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }
    }
}
