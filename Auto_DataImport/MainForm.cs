using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutoItX3Lib;

using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;
using KAutoHelper;
using System.Diagnostics;

namespace Auto_DataImport
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        DataTable data;

        private void btnShowFileDialog_Click(object sender, EventArgs e)
        {
            string fname = "";
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.FileName;
                textBox1.Text = fdlg.FileName;
            }
            else
            {
                return;
            }


            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            data = new DataTable();
            data.Columns.Add(new DataColumn("STT", Type.GetType("System.Int32")));

            // dt.Column = colCount;  
            for (int j = 1; j <= colCount; j++)
            {
                if (xlRange.Cells[1, j] != null && xlRange.Cells[1, j].Value2 != null)
                {
                    data.Columns.Add(xlRange.Cells[1, j].Value2.ToString());
                }
            }
            for (int i = 2; i < rowCount; i++)
            {
                DataRow row = data.NewRow();
                row[0] = i - 1;
                for (int j = 1; j <= colCount; j++)
                {
                    //write the value to the Grid  
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        row[j] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
                data.Rows.Add(row);
            }

            numericUpDown2.Value = data.Rows.Count;
            dataGridView1.DataSource = data;

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:  
            //  never use two dots, all COM objects must be referenced and released individually  
            //  ex: [somthing].[something].[something] is bad  

            //release com objects to fully kill excel process from running in the background  
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release  
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private void btnImportStart_Click(object sender, EventArgs e)
        {
            if (data == null)
            {
                MessageBox.Show("Browse the Excel file to start importing data.");
            }
            else
            {
                var hWnd = IntPtr.Zero;
                var temp = Process.GetProcessesByName("QuanLyPhatTu");
                if (temp.Length == 0)
                {
                    MessageBox.Show("Đăng nhập rồi thử lại.");
                    StartApp();
                }
                else
                {
                    hWnd = temp[0].MainWindowHandle;
                    var clienthWnd = AutoControl.FindHandles(hWnd, "WindowsForms10.MDICLIENT.app.0.265601d_r9_ad1", null);
                    AutoControl.BringToFront(hWnd);

                    for (int i = 0; i < 6; i++)
                    {
                        AutoControl.SendKeyFocus(KeyCode.ESC);
                    }
                    var childhWnds = AutoControl.FindHandles(hWnd, "WindowsForms10.Window.8.app.0.265601d_r9_ad1", null);
                    int buttonIndex = 0;


                    System.Threading.Thread.Sleep(2000);

                    AutoControl.SendClickOnPosition(childhWnds[buttonIndex], 120, 15);
                    AutoControl.SendClickOnPosition(childhWnds[buttonIndex], 50, 50);
                    for (int j = (int)numericUpDown1.Value - 1; j < data.Rows.Count && j < +(int)numericUpDown1.Value + (int)numericUpDown2.Value - 1; j++)
                    {
                        AutoControl.SendKeyFocus(KeyCode.F3);
                        System.Threading.Thread.Sleep(1500);
                        for (int i = 1; i < data.Columns.Count; i++)
                        {
                            if (i != 4)
                            {

                                if (!string.IsNullOrEmpty(data.Rows[j][i].ToString()) && i != 7)
                                {
                                    AutoControl.SendStringFocus(data.Rows[j][i].ToString());
                                }
                                System.Threading.Thread.Sleep(300);
                                if (i == 7)
                                {
                                    if (data.Rows[j][i].ToString().Trim().ToLower() == "nữ")
                                    {
                                        AutoControl.SendKeyFocus(KeyCode.RIGHT);
                                        System.Threading.Thread.Sleep(300);
                                    }
                                }
                                AutoControl.SendKeyFocus(KeyCode.TAB);
                                if (i == 19 || i == 23)
                                {
                                    AutoControl.SendKeyFocus(KeyCode.TAB);
                                }
                            }
                        }
                        AutoControl.SendKeyFocus(KeyCode.ESC);
                        AutoControl.SendKeyFocus(KeyCode.RIGHT);
                        AutoControl.SendKeyFocus(KeyCode.ENTER);
                    }
                }
            }
        }
        private IntPtr StartApp()
        {
            Process p = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = @"C:\VTEK\QLPT\QuanLyPhatTu.exe";
            startInfo.CreateNoWindow = false;
            p.StartInfo = startInfo;
            p.Start();
            return p.Handle;
        }
        private void StartImport(IntPtr[] handles, int[] indexes, DataRow data)
        {
            for (int i = 0; i < indexes.Length; i++)
            {
                AutoControl.SendText(handles[indexes[i]], data[i].ToString());
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            decimal rs = numericUpDown1.Value;
            if (rs < 1)
            {
                numericUpDown1.Value = 1;
            }
            else if (data != null && rs > data.Rows.Count)
            {
                numericUpDown1.Value = data.Rows.Count;
            }
        }
        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            decimal rs = numericUpDown2.Value;
            if (rs < 1)
            {
                numericUpDown2.Value = 1;
            }
            else if (data != null && rs + numericUpDown1.Value > data.Rows.Count + 1)
            {
                numericUpDown2.Value = data.Rows.Count - numericUpDown1.Value + 1;
            }
        }
    }
}
