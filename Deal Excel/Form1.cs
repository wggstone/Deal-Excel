using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;



namespace Deal_Excel
{
    public partial class Form1 : Form
    {
        int BaseSN_c = 4 + 1, BaseDate_c = 6 + 1, BaseOSV_data = 8 + 1, BaseOSV_Info = 9 + 1, BaseOSV_Decision = 11 + 1;
        int Base_Sheet = 1;
        int Out_Sheet = 1;
        //common column only OutOSV_data could less 1
        int OutSN_c = -1, OutDate_c = -1, OutOSV_data = -1, OutOSV_Info = -1, OutOSV_Decision = -1;
        //msi column in weekly report
        int msi_OutSN_c = 9, msi_OutDate_c = 11, msi_OutOSV_data = 12, msi_OutOSV_Info = 16, msi_OutOSV_Decision = 23;
        //lcfc column in weekly report
        int lcfc_OutSN_c = 2, lcfc_OutDate_c = 8, lcfc_OutOSV_data = -1, lcfc_OutOSV_Info = 7, lcfc_OutOSV_Decision = 9;
        //inventec column in weekly report
        int inventec_OutSN_c = 8, inventec_OutDate_c = 9, inventec_OutOSV_data = 3, inventec_OutOSV_Info = 15, inventec_OutOSV_Decision = 17;

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 1) //select usi
            {
                textBox5.Visible = true;
                button7.Visible = true;
            }
            else
            {
                textBox5.Visible = false;
                button7.Visible = false;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "select file: WEEKLY report";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = openFileDialog1.FileName;
            }


        }

        //wistron column in weekly report
        int wistron_OutSN_c = 9, wistron_OutDate_c = 11, wistron_OutOSV_data = 3, wistron_OutOSV_Info = 17, wistron_OutOSV_Decision = 31;
        //USI column in weekly report
        int usi_OutSN_c = 8, usi_OutDate_c = 3, usi_OutOSV_data = 4, usi_OutOSV_Info = 11, usi_OutOSV_Decision = 12;


        BaseData gbasedata = new BaseData();


        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            tabControl1.SelectedIndex = 1;

            // 支持右键拷贝
            ContextMenuStrip listboxMenu = new ContextMenuStrip();
            ToolStripMenuItem copyMenu = new ToolStripMenuItem("Copy");
            copyMenu.Click += new EventHandler(Copy_Click);

            ToolStripMenuItem clearMenu = new ToolStripMenuItem("Clear");
            clearMenu.Click += new EventHandler(Clear_Click);
            listboxMenu.Items.AddRange(new ToolStripItem[] { copyMenu, clearMenu });
            //this.propertyListBox.ContextMenuStrip = listboxMenu;
            listBox1.ContextMenuStrip = listboxMenu; 
        }

        private void Copy_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            
            string tempstr = "";
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                tempstr = tempstr + listBox1.Items[i] + "\x0D\x0A";
            }
            Clipboard.SetDataObject(tempstr);
        }
        private void Clear_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "select file";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "select file";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            char[] spearator = { '/' };
            char[] spearator1 = new char[1];

            if ((textBox1.Text == "") || (textBox2.Text == ""))
            {
                MessageBox.Show(" Please \x0D\x0A select the file", "Confirm Message", MessageBoxButtons.OK);
                return;
            }
            if (textBox1.Text == textBox2.Text)
            {
                MessageBox.Show(" Please do not select the same file", "Confirm Message", MessageBoxButtons.OK);
                return;
            }
            string tempteststr = textBox1.Text;
            string tempteststr1 = textBox2.Text;
            if ((tempteststr.IndexOf("TSC FA Report", StringComparison.OrdinalIgnoreCase) < 0) ||
                (tempteststr1.IndexOf(comboBox1.Items[comboBox1.SelectedIndex].ToString(), StringComparison.OrdinalIgnoreCase) < 0) ||
                (tempteststr1.IndexOf("weekly", StringComparison.OrdinalIgnoreCase) < 0))
            {
                MessageBox.Show(" Please select the right file", "Confirm Message", MessageBoxButtons.OK);
                return;
            }

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;

            int iFaStartLine, iFaEndLine, iWrStartLine, iWrEndLine;
            iFaStartLine = 2; iFaEndLine = -1; iWrStartLine = 2; iWrEndLine = -1;
            iFaStartLine = Convert.ToInt32(FaStartLine.Text);
            iFaEndLine = Convert.ToInt32(FaEndLine.Text);
            iWrStartLine = Convert.ToInt32(WrStartLine.Text);
            iWrEndLine = Convert.ToInt32(WrEndLine.Text);

            // c#中，十六进制 用 \xAE....
            if (MessageBox.Show(" Please check your settings:\x0D\x0A sheet selection,column settings \x0D\x0A start to process?", "Confirm Message", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
            {
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Enabled = true;
                return;
            }

            Base_Sheet = comboBox1.SelectedIndex+1;

            //msi colum
            //int msi_OutSN_c = 9, msi_OutDate_c = 11, msi_OutOSV_data = 12, msi_OutOSV_Info = 16, msi_OutOSV_Decision = 23;
            //lcfc colum
            //int lcfc_OutSN_c = 2, lcfc_OutDate_c = 8, lcfc_OutOSV_data = -1, lcfc_OutOSV_Info = 7, lcfc_OutOSV_Decision = 9;
            if (Base_Sheet == 1) //Msi
            {
                OutSN_c = msi_OutSN_c;
                OutDate_c = msi_OutDate_c;
                OutOSV_data = msi_OutOSV_data;
                OutOSV_Info = msi_OutOSV_Info;
                OutOSV_Decision = msi_OutOSV_Decision;

                Out_Sheet = 2;
                spearator1[0] = '-';
            }
            else if (Base_Sheet == 2) //Usi
            {
                OutSN_c = usi_OutSN_c;
                OutDate_c = usi_OutDate_c;
                OutOSV_data = usi_OutOSV_data;
                OutOSV_Info = usi_OutOSV_Info;
                OutOSV_Decision = usi_OutOSV_Decision;

                Out_Sheet = 3;
                spearator1[0] = '/';
            }
            else if (Base_Sheet == 3) //Wistron
            {
                OutSN_c = wistron_OutSN_c;
                OutDate_c = wistron_OutDate_c;
                OutOSV_data = wistron_OutOSV_data;
                OutOSV_Info = wistron_OutOSV_Info;
                OutOSV_Decision = wistron_OutOSV_Decision;

                Out_Sheet = 1;
                spearator1[0] = '/';
            }
            else if (Base_Sheet == 4) //Wistron AIO
            {
                MessageBox.Show("there are no foundation data for USI ", "Confirm Message", MessageBoxButtons.OKCancel);
                return;
            }
            else if (Base_Sheet == 5) //Inventec
            {
                OutSN_c = inventec_OutSN_c;
                OutDate_c = inventec_OutDate_c;
                OutOSV_data = inventec_OutOSV_data;
                OutOSV_Info = inventec_OutOSV_Info;
                OutOSV_Decision = inventec_OutOSV_Decision;

                Out_Sheet = 2;
                spearator1[0] = '/';
            }
            else if (Base_Sheet == 6) //LCFC
            {
                OutSN_c = lcfc_OutSN_c;
                OutDate_c = lcfc_OutDate_c;
                OutOSV_data = lcfc_OutOSV_data;
                OutOSV_Info = lcfc_OutOSV_Info;
                OutOSV_Decision = lcfc_OutOSV_Decision;

                Out_Sheet = 1;
                spearator1[0] = '/';
            }
            else
            {
                MessageBox.Show("invalid data selected ", "Confirm Message", MessageBoxButtons.OKCancel);
                return;
            }



            listBox1.Items.Clear();

            listBox1.Items.Add(comboBox1.Text);
            listBox1.Items.Add("source: " + textBox2.Text);
            listBox1.Items.Add("dist: " + textBox1.Text);

            object Nothing = System.Reflection.Missing.Value;

            Excel.Application excel = null;
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range1 = null;
            int startrow=1;
            string sn, sndate;

            Excel.Application excel1 = null;
            Excel.Workbooks wbs1 = null;
            Excel.Workbook wb1 = null;
            Excel.Worksheet ws1 = null;
            int startrow1=1;
            string sn1, sndate1;

            try
            {
                excel = new Excel.Application();
                excel.UserControl = true;
                excel.DisplayAlerts = false;

                excel.Application.Workbooks.Open(textBox1.Text, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);

                wbs = excel.Workbooks;
                wb = wbs[1];

            }
            catch (Exception mye)
            {

                MessageBox.Show(mye.ToString() + "  open " + textBox1.Text + "failed");
                return;
            }
            try
            {
                excel1 = new Excel.Application();
                excel1.UserControl = true;
                excel1.DisplayAlerts = false;

                excel1.Application.Workbooks.Open(textBox2.Text, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);

                wbs1 = excel1.Workbooks;
                wb1 = wbs1[1];

            }
            catch (Exception mye)
            {

                MessageBox.Show(mye.ToString() + "  open " + textBox2.Text + "failed");
                return;
            }
            ws = (Excel.Worksheet)wb.Worksheets[Base_Sheet];//Fa report

            ws1 = (Excel.Worksheet)wb1.Worksheets[Out_Sheet];//weekly report
            int totalrows1 = ws1.UsedRange.Rows.Count;
            int totalrows = ws.UsedRange.Rows.Count;

            //iFaStartLine, iFaEndLine, iWrStartLine, iWrEndLine
            if((iFaEndLine<totalrows)&&(iFaEndLine>1))
                totalrows = iFaEndLine;

            if((iWrEndLine<totalrows1)&&(iWrEndLine>1))
                totalrows1 = iWrEndLine;
            
            if(iFaStartLine<2)
                iFaStartLine = 2;

            if(iWrStartLine<2)
                iWrStartLine = 2;

            if ((iFaStartLine > totalrows) || (iWrStartLine > totalrows1) )
            {
                wb.Close();
                excel.Application.Workbooks.Close();
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                excel = null;

                wb1.Close();
                excel1.Application.Workbooks.Close();
                excel1.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1);
                excel1 = null;

                MessageBox.Show("invalid data selected ", "Confirm Message", MessageBoxButtons.OKCancel);
                return;
            }

            string tempstr, tempstr1;
            String[] strlist, strlist1;
            int year, month, day, year1, month1, day1;

            tempstr1 = ((Excel.Range)ws1.Cells[104, 11]).Text;
            tempstr1 = tempstr1.Trim();
            strlist1 = tempstr1.Split(spearator1);
            /*
            //begin: test for date format: the cell must be legal date format yyyy/mm/dd otherwise the method willDateTime.FromOADate throw an ArgumentException
            Excel.Range rng;
            DateTime date2;
            
            rng = (Excel.Range)ws.Cells[2, 12];
            tempstr = rng.Text;
            tempstr1 = rng.Value2.ToString();
            try
            {
                date2 = DateTime.FromOADate(double.Parse(tempstr1));
            }
            catch (ArgumentException mye)
            {
                Console.Write("Exception Thrown: ");
                Console.Write("{0}", mye.GetType(), mye.Message);            
            }
            //end: test for date format
            */

            List<Identifier> ilist = new List<Identifier>();
            
            for (int i = iFaStartLine; i <= totalrows; i++)
            {
                Identifier tempidentifier = new Identifier();
                tempstr = ((Excel.Range)ws.Cells[i, BaseSN_c]).Text;
                tempstr = tempstr.Trim();
                tempidentifier.sn = tempstr.ToUpper();
                tempstr = ((Excel.Range)ws.Cells[i, BaseDate_c]).Text;
                tempstr = tempstr.Trim();
                strlist = tempstr.Split(spearator);
                tempidentifier.lineno = i;
                if (strlist.Length == 3)
                {
                    tempidentifier.year = Convert.ToInt32(strlist[0]);
                    tempidentifier.month = Convert.ToInt32(strlist[1]);
                    tempidentifier.day = Convert.ToInt32(strlist[2]);
                }
                else
                {
                    tempidentifier.year = -1;
                    tempidentifier.month = 0;
                    tempidentifier.day = 0;
                    listBox1.Items.Add("file 1,line " + i + " " + tempidentifier.sn + " wrong date format");
                }
                ilist.Add(tempidentifier);
            }
            int ii = ilist.Count;
            string tempsn;
            int j = 0;
            Excel.Range temprange = null;
            for (int i = iWrStartLine; i <= totalrows1; i++)
            {

                tempsn = ((Excel.Range)ws1.Cells[i, OutSN_c]).Text;
                tempsn = tempsn.Trim();
                tempsn = tempsn.ToUpper();
                tempstr = ((Excel.Range)ws1.Cells[i, OutDate_c]).Text;
                tempstr = tempstr.Trim();
                strlist = tempstr.Split(spearator1);
                if(i==787)
                {
                    int tempi = 0;
                }

                if (strlist.Length == 3)
                {
                    if (Base_Sheet == 2)//usi
                    {
                        year = Convert.ToInt32(strlist[0]);
                        month = Convert.ToInt32(strlist[1]);
                        day = Convert.ToInt32(strlist[2]);
                    }
                    else
                    {
                        year = Convert.ToInt32(strlist[0]);
                        if (Base_Sheet == 3 && year < 23) //Wistron
                            year = year + 2000;
                        month = Convert.ToInt32(strlist[1]);
                        day = Convert.ToInt32(strlist[2]);
                    }
                    
                    for (j = 0; j < ilist.Count; j++)
                    {

                        if (tempsn == ilist[j].sn && year == ilist[j].year && month == ilist[j].month && day == ilist[j].day)
                        {
                            //int BaseSN_c = 4, BaseDate_c = 6, BaseOSV_data = 8, BaseOSV_Info = 9, BaseOSV_Decision = 11;
                            //int OutSN_c = 9, OutDate_c = 11, OutOSV_data = 12, OutOSV_Info = 16, OutOSV_Decision = 23;
                            if (OutOSV_data > 0)
                            {
                                tempstr = ((Excel.Range)ws1.Cells[i, OutOSV_data]).Text;
                                ws.Cells[ilist[j].lineno, BaseOSV_data] = tempstr;
                            }
                            if (OutOSV_Info > 0)
                            {
                                tempstr = ((Excel.Range)ws1.Cells[i, OutOSV_Info]).Text;
                                ws.Cells[ilist[j].lineno, BaseOSV_Info] = tempstr;
                            }
                            if (OutOSV_Decision > 0)
                            {
                                tempstr = ((Excel.Range)ws1.Cells[i, OutOSV_Decision]).Text;
                                ws.Cells[ilist[j].lineno, BaseOSV_Decision] = tempstr;
                            }
                            temprange = ws.Range[ilist[j].lineno.ToString() + ":" + ilist[j].lineno.ToString()];
                            temprange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            break;
                        }
                    }
                    if (j == ilist.Count)
                    {
                        listBox1.Items.Add("file 2,line " + i + " " + tempsn + " couldn't find item in FA file");
                    }
                    /*else
                    {
                        listBox1.Items.Add("file 1,line " + (j+2) + " " + tempsn + " added the info to  FA file sucessfully!");
                        break;
                    }*/
                }
                else
                {
                    year = 0;
                    month = 0;
                    day = 0;
                    listBox1.Items.Add("file 2,line " + i + " " + tempsn + " wrong date format");
                }
            }



/*
            for (int i = 1; i <= totalrows1; i++)
            {
                ws1.Cells[2, 7] = "wggtest";
                for (int j = startrow; j < totalrows; j++)
                {
                    sn1 = ((Excel.Range)ws1.Cells[i, OutSN_c]).Text;
                    sn = ((Excel.Range)ws.Cells[j,BaseSN_c]).Text;
                    sndate1 = ((Excel.Range)ws1.Cells[i, OutDate_c]).Text;
                    sndate = ((Excel.Range)ws.Cells[j,BaseDate_c]).Text;
                    sn = sn.Trim();
                    sn1 = sn1.Trim();
                    sndate1 = sndate1.Trim();
                    sndate = sndate.Trim();

                    if (sn == sn1 && sndate == sndate1)
                    {
                        break;
                    }
                }
            
            }
*/
            wb.Save();
            //wb1.Save();

            //finally
            {
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Enabled = true;

                if (excel != null)
                {
                    if (wbs != null)
                    {
                        if (wb != null)
                        {
                            if (ws != null)
                            {
                                if (range1 != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range1);
                                    range1 = null;
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                                ws = null;
                            }
                            wb.Close(false, Nothing, Nothing);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                            wb = null;
                        }
                        wbs.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wbs);
                        wbs = null;
                    }
                    excel.Application.Workbooks.Close();
                    excel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                    excel = null;
                }
                
                if (excel1 != null)
                {
                    if (wbs1 != null)
                    {
                        if (wb1 != null)
                        {
                            if (ws1 != null)
                            {
                                if (range1 != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range1);
                                    range1 = null;
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws1);
                                ws1 = null;
                            }
                            wb1.Close(false, Nothing, Nothing);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb1);
                            wb1 = null;
                        }
                        wbs1.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wbs1);
                        wbs1 = null;
                    }
                    excel1.Application.Workbooks.Close();
                    excel1.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1);
                    excel1 = null;
                }
                GC.Collect();
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "select file: WEEKLY report";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = openFileDialog1.FileName;
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "select file: DAILY report";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = openFileDialog1.FileName;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int customerID;

            button6.Enabled = false;
            button7.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;

            DateTime DailyReportDate = dateTimePicker1.Value;

            customerID = comboBox2.SelectedIndex;

            if ((customerID == 3))//wistron aio
            {
                MessageBox.Show(" Until now \x0D\x0A " + comboBox2.Items[comboBox2.SelectedIndex].ToString() + "\x0D\x0A not supported", "Confirm Message", MessageBoxButtons.OK);
                return;
            }
            if (customerID == 1)//usi
            {
                if (textBox5.Text == "")
                {
                    MessageBox.Show("Please select usi special file", "Confirm Message", MessageBoxButtons.OK);
                    return;
                }
            }
            if ((textBox3.Text == "") || (textBox4.Text == ""))
            {
                MessageBox.Show(" Please \x0D\x0A select the file", "Confirm Message", MessageBoxButtons.OK);
                return;
            }
            if (textBox3.Text == textBox4.Text)
            {
                MessageBox.Show(" Please do not select the same file", "Confirm Message", MessageBoxButtons.OK);
                return;
            }
            string tempteststr = textBox3.Text;
            string tempteststr1 = textBox4.Text;
            if ((tempteststr.IndexOf(comboBox2.Items[customerID].ToString(), StringComparison.OrdinalIgnoreCase) < 0) ||
                (tempteststr.IndexOf("weekly", StringComparison.OrdinalIgnoreCase) < 0) ||
                (tempteststr1.IndexOf(comboBox2.Items[customerID].ToString(), StringComparison.OrdinalIgnoreCase) < 0))
            //(tempteststr1.IndexOf("daily", StringComparison.OrdinalIgnoreCase) < 0))
            {
                MessageBox.Show(" Please select the right file", "Confirm Message", MessageBoxButtons.OK);
                return;
            }
            if (!(tempteststr1.Contains(DailyReportDate.Year.ToString()) && tempteststr1.Contains(DailyReportDate.Month.ToString()) && tempteststr1.Contains(DailyReportDate.Day.ToString())))
            {
                MessageBox.Show(" Please select the right Daily Report Date", "Confirm Message", MessageBoxButtons.OK);
                return;
            }



            gbasedata.Init(customerID);

            //listBox1.Items.Clear();

            object Nothing = System.Reflection.Missing.Value;

            Excel.Application excel = null;
            Excel.Workbooks wbs = null;
            Excel.Workbook wbWeekly = null;
            Excel.Worksheet wsWeekly = null;
            Excel.Workbook wbDaily = null;
            Excel.Worksheet wsDaily = null;
            Excel.Workbook wbUsiSpecialWeekly = null;
            Excel.Worksheet wsUsiSpecialWeekly = null;
            Excel.Range range1 = null;
            Excel.Range range2 = null;

            int TotalRowsWeekly = 0;
            int TotalRowsDaily = 0;
            int TotalRowsUsiSpecialWeekly = 0;
            try
            {
                excel = new Excel.Application();
                excel.UserControl = true;
                excel.DisplayAlerts = false;

                //excel.Application.Workbooks.Open(textBox1.Text, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
                wbs = excel.Workbooks;
                wbWeekly = wbs.Open(textBox3.Text, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
                wbDaily = wbs.Add(textBox4.Text);
                if (customerID == 1) //usi
                {
                    wbUsiSpecialWeekly = wbs.Open(textBox5.Text);
                    wsUsiSpecialWeekly = wbUsiSpecialWeekly.Sheets["BUDAPEST"];
                }

                wsWeekly = wbWeekly.Sheets[gbasedata.SheetID_Weekly];
                wsDaily = wbDaily.Sheets[1];


            }
            catch (Exception mye)
            {

                MessageBox.Show(mye.ToString() + "  open " + textBox1.Text + "failed");
                return;
            }

            TotalRowsWeekly = wsWeekly.UsedRange.Rows.Count;
            TotalRowsDaily = wsDaily.UsedRange.Rows.Count;
            if (customerID == 1) //usi
                TotalRowsUsiSpecialWeekly = wsUsiSpecialWeekly.UsedRange.Rows.Count;
            /*
            //bengin:demo how to select the range and how to insert,cut line
            //Range sourceRange = ws1.Range["L5:L37"];
            //Range destinationRange = ws1.Range["M5"];
            //sourceRange.Cut(destinationRange);

            range1 = wsWeekly.Range["2:2"];
            range1.Insert(XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            range2 = wsWeekly.Range["4:4"];
            range1 = wsWeekly.Range["2:2"];
            range2.Cut(range1);

            range1 = wsWeekly.Range["2:2"];//对range进行操作(cut,copy...)后，range可能失效。需要重新定义range，否则直接用原来的range会有异常。
            //range1.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(220, 20, 60));
            range1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //end:demo how to select the range and how to insert,cut line
            */
            string tempstr, tempstr1;
            int startline = TotalRowsWeekly, endline = TotalRowsWeekly;
            int i, j, k;
            k = TotalRowsUsiSpecialWeekly + 1;
            for (i = TotalRowsWeekly; i >= 1; i--)
            {
                tempstr = ((Excel.Range)wsWeekly.Cells[i, gbasedata.SN_Weekly]).Text;
                tempstr1 = ((Excel.Range)wsWeekly.Cells[i, gbasedata.Decision_Weekly]).Text;
                tempstr = tempstr.Trim();
                tempstr1 = tempstr1.Trim();
                if (tempstr == "")
                {
                    listBox1.Items.Add("Weekly report: Line " + i + " SN is empty!");
                    goto finalwork;
                }
                if (tempstr1 != "")
                {
                    startline = i + 1;
                    break;
                }
            }
            int usispecialadded = 0;
            int normallyadded = 0;
            int wistronRTV = 0;
            for (i = 2; i <= TotalRowsDaily; i++)
            {
                tempstr = ((Excel.Range)wsDaily.Cells[i, gbasedata.SN_Daily]).Text;
                tempstr = tempstr.Trim();
                if (tempstr == "")
                {
                    listBox1.Items.Add("Daily report: Line " + i + " SN is empty!");
                    break;
                }
                for (j = startline; j <= wsWeekly.UsedRange.Rows.Count; j++)
                {
                    tempstr1 = ((Excel.Range)wsWeekly.Cells[j, gbasedata.SN_Weekly]).Text;
                    tempstr1 = tempstr1.Trim();
                    if (tempstr == tempstr1)
                    {
                        break;
                    }
                }
                //begin: seperate SB27B26156/STA7B26151/STA7B26153/STA7B26155 from usi weekly report.
                if (customerID == 1)
                {
                    tempstr = ((Excel.Range)wsDaily.Cells[i, gbasedata.PN_Daily]).Text;
                    tempstr = tempstr.Replace(" ", "");
                    tempstr = tempstr.ToUpper();
                    if (tempstr == "SB27B26156" || tempstr == "STA7B26151" || tempstr == "STA7B26153" || tempstr == "STA7B26155" || tempstr == "STA7B26152" || tempstr == "STA7B26154" || tempstr == "SB27B75942")
                    {
                        //add to other file
                        FillUsiSpecialReport(wsDaily, i, wsUsiSpecialWeekly, k, customerID, gbasedata, DailyReportDate);
                        k++;
                        usispecialadded++;
                        continue;
                    }
                }
                //end: seperate SB27B26156/STA7B26151/STA7B26153/STA7B26155 from usi weekly report.

                //when add the line at the end of the file, insert a line to obtain the format of the line.
                if (j == (wsWeekly.UsedRange.Rows.Count + 1))
                {
                    range1 = wsWeekly.Range[j.ToString() + ":" + j.ToString()];
                    range1.Insert(XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                }
                FillWeeklyReport(wsDaily, i, wsWeekly, j, customerID, gbasedata, DailyReportDate);
                if (customerID == 2)//wistron
                {
                    if ("SPQL" == wsWeekly.Cells[j, gbasedata.Decision_Weekly].Text)
                    {
                        wistronRTV++;
                    }
                }
                normallyadded++;
                if (j != startline)//only need insert and cut operation when j!= startline
                {
                    //cut and inset
                    range1 = wsWeekly.Range[startline.ToString() + ":" + startline.ToString()];
                    range1.Insert(XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    j++;
                    range2 = wsWeekly.Range[j.ToString() + ":" + j.ToString()];
                    range1 = wsWeekly.Range[startline.ToString() + ":" + startline.ToString()];
                    range2.Cut(range1);

                    //cut method just cut the data, but dosen't delete the line. So, need to delete the line. 
                    range2 = wsWeekly.Range[j.ToString() + ":" + j.ToString()];
                    range2.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                }
                //set range1's border
                range1 = wsWeekly.Range["A" + startline.ToString() + ":" + "AT" + startline.ToString()];
                //range1.BorderAround(XlLineStyle.xlContinuous); //only set the range's border
                range1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range1.Borders.Weight = Excel.XlBorderWeight.xlThin;
                range1.Borders.Color = Color.Black;

                //set the seriel No.
                if ((customerID != 0) && (customerID != 5))//MSI&LCFC
                {
                    wsWeekly.Cells[startline, 1] = Convert.ToString(startline - 1);
                }

                //move startline 
                startline++;
            }
            if ((customerID != 0) && (customerID != 5))//MSI&LCFC
            {
                for (i = startline; i <= wsWeekly.UsedRange.Rows.Count; i++)
                {
                    wsWeekly.Cells[i, 1] = Convert.ToString(i - 1);
                }
            }
            if ((customerID == 1))
            {
                listBox1.Items.Add(comboBox2.Text + " " + dateTimePicker1.Text + ": " + "normal usi  " + normallyadded.ToString() + " Lines added");
                listBox1.Items.Add(comboBox2.Text + " " + dateTimePicker1.Text + ": " + "special usi " + usispecialadded.ToString() + " Lines added");
            }
            else
            {
                listBox1.Items.Add(comboBox2.Text + " " + dateTimePicker1.Text + ": " + normallyadded.ToString() + " Lines added");
                if(customerID==2)
                    listBox1.Items.Add(comboBox2.Text + " " + dateTimePicker1.Text + " RTV: " + wistronRTV.ToString() + " Lines added");
            }
            wbWeekly.Save();
            if (customerID == 1)
                wbUsiSpecialWeekly.Save();
//release the resource of the excel component
finalwork:  
            button4.Enabled = true;
            button5.Enabled = true;
            button7.Enabled = true;
            if (excel != null)
            {
                if (wbs != null)
                {
                    if (wbWeekly != null)
                    {
                        if (wsWeekly != null)
                        {
                            if (range1 != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(range1);
                                range1 = null;
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wsWeekly);
                            wsWeekly = null;
                        }
                        wbWeekly.Close(false, Nothing, Nothing);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wbWeekly);
                        wbWeekly = null;
                    }
                    if (wbDaily != null)
                    {
                        if (wsDaily != null)
                        {
                            if (range1 != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(range1);
                                range1 = null;
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wsDaily);
                            wsDaily = null;
                        }
                        wbDaily.Close(false, Nothing, Nothing);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wbDaily);
                        wbDaily = null;
                    }
                    if (wbUsiSpecialWeekly != null)
                    {
                        if (wsUsiSpecialWeekly != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wsUsiSpecialWeekly);
                            wsUsiSpecialWeekly = null;
                        }
                        wbUsiSpecialWeekly.Close(false, Nothing, Nothing);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wbUsiSpecialWeekly);
                        wbUsiSpecialWeekly = null;
                    }
                    wbs.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wbs);
                    wbs = null;
                }
                excel.Application.Workbooks.Close();
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                excel = null;
            }

        }

        private bool FillUsiSpecialReport(Excel.Worksheet wsDaily, int wsDailyLine, Excel.Worksheet wsWeekly, int wsWeeklyLine, int custermorID, BaseData bassedata, DateTime DailyReportDate)
        {
            wsWeekly.Cells[wsWeeklyLine, 1] = "USI Guad";
            wsWeekly.Cells[wsWeeklyLine, 2] = "BUDAPEST";
            wsWeekly.Cells[wsWeeklyLine, 3] = wsDaily.Cells[wsDailyLine, bassedata.FailDate_Daily].Text; ; //lenovo test date
            wsWeekly.Cells[wsWeeklyLine, 4] = wsDaily.Cells[wsDailyLine, bassedata.PN_Daily].Text;
            wsWeekly.Cells[wsWeeklyLine, 5] = wsDaily.Cells[wsDailyLine, bassedata.SN_Daily].Text;
            wsWeekly.Cells[wsWeeklyLine, 6] = wsDaily.Cells[wsDailyLine, bassedata.Failure_Daily].Text;
            wsWeekly.Cells[wsWeeklyLine, 10] = wsDaily.Cells[wsDailyLine, bassedata.FaResult_Daily].Text;
            wsWeekly.Cells[wsWeeklyLine, 11] = wsDaily.Cells[wsDailyLine, bassedata.Decision_Daily].Text;

            string tempstr, tempstr1;
            string[] strlist;
            int year, month, day;
            bool dateformatok = true;

            tempstr = wsDaily.Cells[wsDailyLine, bassedata.FailDate_Daily].Text;
            tempstr = tempstr.Trim();
            strlist = tempstr.Split('/');
            if (strlist.Length == 3)
            {
                DateTime dt, dt1;
                dt = DateTime.Now;
                dt1 = DateTime.Now;
                year = Convert.ToInt32(strlist[0]);
                month = Convert.ToInt32(strlist[1]);
                day = Convert.ToInt32(strlist[2]);
                try
                {
                    dt = new DateTime(year, 1, 1);
                    dt1 = new DateTime(year, month, day);
                }
                catch (Exception e)
                {
                    listBox1.Items.Add("Daily report,line " + wsDailyLine + " wrong date format: " + tempstr);
                    dateformatok = false;
                    year = -1;
                    month = 0;
                    day = 0;

                    //return false;
                }
                if (dateformatok)
                {
                    TimeSpan ts = dt1 - dt;
                    int daycount = ts.Days + 1;
                    daycount += Convert.ToInt32(dt.DayOfWeek);

                    //if ((daycount % 7) == 0)
                    //daycount++;
                    int weeknum = (int)Math.Ceiling(daycount / 7F);
                    tempstr1 = "";
                    if (custermorID == 2)//Wistron
                        tempstr1 = "wk" + (weeknum).ToString("D2");//Convert.ToString(weeknum - 1,"D2");
                    else if (custermorID == 1)//USI
                        //tempstr1 = "22" + Convert.ToString(weeknum);
                        tempstr1 = (weeknum).ToString("D2");
                    else if (custermorID != 0)//msi
                        tempstr1 = Convert.ToString(weeknum);
                    wsWeekly.Cells[wsWeeklyLine, 17] = tempstr1;
                    wsWeekly.Cells[wsWeeklyLine, 15] = strlist[0];
                    wsWeekly.Cells[wsWeeklyLine, 16] = strlist[1];
                }
            }
            else
            {
                year = -1;
                month = 0;
                day = 0;
                listBox1.Items.Add("Daily report,line " + wsDailyLine + " wrong date format: " + tempstr);
                //return false;
            }
            wsWeekly.Cells[wsWeeklyLine, 14] = (wsWeeklyLine - 1).ToString();

            tempstr = wsDaily.Cells[wsDailyLine, bassedata.PN_Daily].Text;
            tempstr = tempstr.Trim();
            tempstr = tempstr.ToUpper();
            /*for (int i = 0; i < bassedata.ilist.Count; i++)
            {
                if (tempstr == bassedata.ilist[i].pn)
                {
                    wsWeekly.Cells[wsWeeklyLine, 18] = bassedata.ilist[i].modeName;
                    wsWeekly.Cells[wsWeeklyLine, 19] = bassedata.ilist[i].name;
                    wsWeekly.Cells[wsWeeklyLine, 20] = bassedata.ilist[i].type;
                    break;
                }
            }*/
            if (bassedata.ilist.ContainsKey(tempstr))
            {
                wsWeekly.Cells[wsWeeklyLine, 18] = bassedata.ilist[tempstr].modeName;
                wsWeekly.Cells[wsWeeklyLine, 19] = bassedata.ilist[tempstr].name;
                wsWeekly.Cells[wsWeeklyLine, 20] = bassedata.ilist[tempstr].type;
            }
            return true;
        }
        private bool FillWeeklyReport(Excel.Worksheet wsDaily, int wsDailyLine, Excel.Worksheet wsWeekly, int wsWeeklyLine, int custermorID, BaseData bassedata, DateTime DailyReportDate)
        {
            string tempstr, tempstr1;
            string[] strlist;
            bool dateformatok = true;
            DateTime now = DateTime.Now;
            string datestr = now.Year.ToString() + "-" + now.Month.ToString() + "-" + now.Day.ToString();
            int year, month, day;
            if(custermorID != 5)// not lcfc
                tempstr = wsDaily.Cells[wsDailyLine, bassedata.FailDate_Daily].Text;
            else //lcfc ask for use the 
                tempstr = wsDaily.Cells[wsDailyLine, bassedata.FailDate_Daily+3].Text;
            tempstr = tempstr.Trim();
            strlist = tempstr.Split('/');
            if (strlist.Length == 3)
            {
                DateTime dt, dt1;
                dt = DateTime.Now;
                dt1 = DateTime.Now;
                year = Convert.ToInt32(strlist[0]);
                month = Convert.ToInt32(strlist[1]);
                day = Convert.ToInt32(strlist[2]);
                try
                {
                    dt = new DateTime(year, 1, 1);
                    dt1 = new DateTime(year, month, day);
                }
                catch (Exception e)
                {
                    listBox1.Items.Add("Daily report,line " + wsDailyLine + " wrong date format: " + tempstr);
                    dateformatok = false;
                    year = -1;
                    month = 0;
                    day = 0;

                    //return false;
                }
                if (dateformatok)
                {
                    TimeSpan ts = dt1 - dt;
                    int daycount = ts.Days+1;
                    daycount += Convert.ToInt32(dt.DayOfWeek);

                    //if ((daycount % 7) == 0)
                        //daycount++;
                    int weeknum = (int)Math.Ceiling(daycount / 7F);
                    tempstr1 = "";
                    if (custermorID == 2)//Wistron
                        tempstr1 = "wk" + (weeknum).ToString("D2");//Convert.ToString(weeknum - 1,"D2");
                    else if (custermorID == 1)//USI
                        //tempstr1 = "22" + Convert.ToString(weeknum);
                        tempstr1 = strlist[0].Substring(2,2) + (weeknum).ToString("D2");
                    else if (custermorID != 0)//msi
                        tempstr1 = Convert.ToString(weeknum);
                    if (bassedata.Weeknum_Weekly != -1)
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Weeknum_Weekly] = tempstr1;
                }
            }
            else
            {
                year = -1;
                month = 0;
                day = 0;
                listBox1.Items.Add("Daily report,line " + wsDailyLine + " wrong date format: " + tempstr);
                //return false;
            }

            wsWeekly.Cells[wsWeeklyLine, bassedata.PN_Weekly] = wsDaily.Cells[wsDailyLine, bassedata.PN_Daily].Text;
            wsWeekly.Cells[wsWeeklyLine, bassedata.SN_Weekly] =   wsDaily.Cells[wsDailyLine, bassedata.SN_Daily].Text;
            wsWeekly.Cells[wsWeeklyLine, bassedata.FailDate_Weekly] = wsDaily.Cells[wsDailyLine, bassedata.FailDate_Daily].Text;
            wsWeekly.Cells[wsWeeklyLine, bassedata.Failure_Weekly] = wsDaily.Cells[wsDailyLine, bassedata.Failure_Daily].Text;
            wsWeekly.Cells[wsWeeklyLine, bassedata.FaResult_Weekly] = wsDaily.Cells[wsDailyLine, bassedata.FaResult_Daily].Text;
            
            if (custermorID == 0)//MSI
            {
                if (bassedata.DailyReportDate_Weekly != -1)
                {
                    wsWeekly.Cells[wsWeeklyLine, bassedata.DailyReportDate_Weekly] = DailyReportDate.Year.ToString() + "-" + DailyReportDate.Month.ToString("D2") + "-" + DailyReportDate.Day.ToString("D2");
                }
                wsWeekly.Cells[wsWeeklyLine, bassedata.Decision_Weekly] = wsDaily.Cells[wsDailyLine, bassedata.Decision_Daily].Text;
                wsWeekly.Cells[wsWeeklyLine, 1] = "TSC";
                wsWeekly.Cells[wsWeeklyLine, 3] = "Lenovo";
                wsWeekly.Cells[wsWeeklyLine, 4] = "Hungary";
                wsWeekly.Cells[wsWeeklyLine, 5] = "MB";
                wsWeekly.Cells[wsWeeklyLine, 10] = wsWeekly.Cells[wsWeeklyLine, bassedata.SN_Weekly];
                wsWeekly.Cells[wsWeeklyLine, 12] = datestr;
                if(year !=-1)
                {
                    tempstr = year.ToString() + "-" + month.ToString() + "-" + day.ToString();
                    wsWeekly.Cells[wsWeeklyLine, bassedata.FailDate_Weekly] = tempstr;
                }
            }

            if (custermorID == 1)//USI
            {
                if (bassedata.DailyReportDate_Weekly != -1)
                {
                    wsWeekly.Cells[wsWeeklyLine, bassedata.DailyReportDate_Weekly] = DailyReportDate.Year.ToString() + "/" + DailyReportDate.Month.ToString("D2") + "/" + DailyReportDate.Day.ToString("D2");
                }
                wsWeekly.Cells[wsWeeklyLine, bassedata.Decision_Weekly] = wsDaily.Cells[wsDailyLine, bassedata.Decision_Daily].Text;
                wsWeekly.Cells[wsWeeklyLine, 4] = "HGY";
                tempstr = wsWeekly.Cells[wsWeeklyLine, bassedata.PN_Weekly].Text;
                tempstr = tempstr.Trim();
                tempstr = tempstr.ToUpper();
                /*for (int i = 0; i < bassedata.ilist.Count; i++)
                {
                    if (tempstr == bassedata.ilist[i].pn)
                    {
                        wsWeekly.Cells[wsWeeklyLine, 6] = bassedata.ilist[i].type;
                        wsWeekly.Cells[wsWeeklyLine, 7] = bassedata.ilist[i].name;
                        break;
                    }
                }*/
                if (bassedata.ilist.ContainsKey(tempstr))
                {
                    wsWeekly.Cells[wsWeeklyLine, 6] = bassedata.ilist[tempstr].type;
                    wsWeekly.Cells[wsWeeklyLine, 7] = bassedata.ilist[tempstr].name;
                }

                //fill columns relate to decision
                /*tempstr = wsDaily.Cells[wsDailyLine, bassedata.Decision_Daily].Text;
                tempstr = tempstr.Trim();
                switch (tempstr)
                {
                    case "CID":
                        wsWeekly.Cells[wsWeeklyLine, bassedata.CidBitmap_Weekly] = "1";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "SCRAPT";
                        break;
                    case "RTV":
                        wsWeekly.Cells[wsWeeklyLine, bassedata.RtvBitmap_Weekly] = "1";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "RMA";
                        break;
                    case "NDF":
                    case "RTL":
                        wsWeekly.Cells[wsWeeklyLine, bassedata.NdfBitmap_Weekly] = "1";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "RTL";
                        break;
                    //case "FW":
                    //    wsWeekly.Cells[wsWeeklyLine, bassedata.FwfBitmap_Weekly] = "1";
                    //    wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "RTL";
                    //    break;
                    default:
                        listBox1.Items.Add("Daily report: Line " + wsDailyLine + " Decision is illegel: " + tempstr);
                        break;
                }*/
            }

            if (custermorID == 4)//Inventec
            {
                if (bassedata.DailyReportDate_Weekly != -1)
                {
                    wsWeekly.Cells[wsWeeklyLine, bassedata.DailyReportDate_Weekly] = DailyReportDate.Year.ToString() + "/" + DailyReportDate.Month.ToString("D2") + "/" + DailyReportDate.Day.ToString("D2");
                }
                wsWeekly.Cells[wsWeeklyLine, 4] = wsDaily.Cells[wsDailyLine, 4].Text;// material type
                wsWeekly.Cells[wsWeeklyLine, 5] = "INVENTEC";
                wsWeekly.Cells[wsWeeklyLine, 11] = "Production line function test station";
                wsWeekly.Cells[wsWeeklyLine, 13] = "1/A";
                wsWeekly.Cells[wsWeeklyLine, 14] = wsDaily.Cells[wsDailyLine, 14].Text;//supplier SN
                wsWeekly.Cells[wsWeeklyLine, bassedata.Decision_Weekly] = wsDaily.Cells[wsDailyLine, bassedata.Decision_Daily].Text;
                //fill columns relate to decision
                tempstr = wsDaily.Cells[wsDailyLine, bassedata.Decision_Daily].Text;
                tempstr = tempstr.Trim();
                tempstr = tempstr.ToUpper();
                switch (tempstr)
                {
                    case "DAMAGE":
                        tempstr1 = wsWeekly.Cells[wsWeeklyLine, bassedata.FaResult_Weekly].Text;
                        tempstr1 = tempstr1.Replace(" ", "");
                        tempstr1 = tempstr1.ToLower();
                        if(tempstr1.IndexOf("unrepairable",StringComparison.OrdinalIgnoreCase)<0 &&
                           tempstr1.IndexOf("notrepairable", StringComparison.OrdinalIgnoreCase)< 0)
                            wsWeekly.Cells[wsWeeklyLine, 17] = "RFR";
                        else
                            wsWeekly.Cells[wsWeeklyLine, 17] = "Scrap";
                        wsWeekly.Cells[wsWeeklyLine, 18] = "5/A";
                        wsWeekly.Cells[wsWeeklyLine, 20] = "1";
                        break;
                    case "SPQL":
                        wsWeekly.Cells[wsWeeklyLine, 17] = "RTV";
                        wsWeekly.Cells[wsWeeklyLine, 18] = "1/A";
                        wsWeekly.Cells[wsWeeklyLine, 19] = "1";
                        break;
                    case "NDF":
                    case "RTL":
                        wsWeekly.Cells[wsWeeklyLine, 17] = "RTL";
                        wsWeekly.Cells[wsWeeklyLine, 18] = "5/A";
                        wsWeekly.Cells[wsWeeklyLine, 21] = "1";
                        break;
                    case "FW":
                        wsWeekly.Cells[wsWeeklyLine, 17] = "RTL";
                        wsWeekly.Cells[wsWeeklyLine, 18] = "5/A";
                        wsWeekly.Cells[wsWeeklyLine, 22] = "1";
                        break;
                    default:
                        listBox1.Items.Add("Daily report: Line " + wsDailyLine + " Decision is illegel: " + tempstr);
                        break;
                }

                //fill product name column weekly report
                tempstr = wsWeekly.Cells[wsWeeklyLine, bassedata.PN_Weekly].Text;
                tempstr = tempstr.Trim();
                tempstr = tempstr.ToUpper();
                /*for (int i = 0; i < bassedata.ilist.Count; i++)
                {
                    if (tempstr == bassedata.ilist[i].pn)
                    {
                        wsWeekly.Cells[wsWeeklyLine, 6] = bassedata.ilist[i].name;//product name
                        wsWeekly.Cells[wsWeeklyLine, 12] = bassedata.ilist[i].type;
                        break;
                    }
                }*/
                if (bassedata.ilist.ContainsKey(tempstr))
                {
                    wsWeekly.Cells[wsWeeklyLine, 6] = bassedata.ilist[tempstr].name;//product name
                    wsWeekly.Cells[wsWeeklyLine, 12] = bassedata.ilist[tempstr].type;
                }

            }

            if (custermorID == 2)//Wistron
            {
                if (bassedata.DailyReportDate_Weekly != -1)
                {
                    wsWeekly.Cells[wsWeeklyLine, bassedata.DailyReportDate_Weekly] = DailyReportDate.Year.ToString() + "/" + DailyReportDate.Month.ToString("D2") + "/" + DailyReportDate.Day.ToString("D2");
                }
                wsWeekly.Cells[wsWeeklyLine, bassedata.SN_Weekly + 7] = wsDaily.Cells[wsDailyLine, bassedata.SN_Daily].Text;
                tempstr = wsDaily.Cells[wsDailyLine, bassedata.Decision_Daily].Text;
                tempstr = tempstr.Trim();
                switch(tempstr)
                {
                    case "CID":
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Decision_Weekly] = "Damage";
                        tempstr1 = wsWeekly.Cells[wsWeeklyLine, bassedata.FaResult_Weekly].Text;
                        tempstr1 = tempstr1.Replace(" ", "");
                        tempstr1 = tempstr1.ToLower();
                        if (tempstr1.IndexOf("unrepairable",StringComparison.OrdinalIgnoreCase)<0 &&
                           tempstr1.IndexOf("notrepairable", StringComparison.OrdinalIgnoreCase)< 0)
                            wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "CID";
                        else
                            wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "Scrap";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.CidBitmap_Weekly] = "1";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Responsibility_Weekly] = "CID.客戶責任";
                        break;
                    case "RTV":
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Decision_Weekly] = "SPQL";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "RTV FA";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.RtvBitmap_Weekly] = "1";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Responsibility_Weekly] = "VID.廠商責任";
                        break;
                    case "NDF":
                    case "RTL":  //wistron 报告中应该没有 RTL，RTL也算NDF
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Decision_Weekly] = "NDF";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "RTL";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.NdfBitmap_Weekly] = "1";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Responsibility_Weekly] = "NDF.誤判";
                        break;
                    case "FW":
                        //2023.4.22 wistrong ask for change "转好品" to "return to PD"
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Decision_Weekly] = "return to PD";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Disposition_Weekly] = "RTL";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.FwfBitmap_Weekly] = "1";
                        wsWeekly.Cells[wsWeeklyLine, bassedata.Responsibility_Weekly] = "CID.客戶責任";
                        break;
                    default:
                        listBox1.Items.Add("Daily report: Line " + wsDailyLine + " Decision is illegel: " + tempstr);
                        break;
                }
                tempstr = wsDaily.Cells[wsDailyLine, bassedata.PN_Daily].Text;
                tempstr = tempstr.Trim();
                tempstr = tempstr.ToUpper();
                /*if (tempstr == "SC57A02007")
                {
                    tempstr = "SC57A02007";
                }*/
                /*for (int i = 0; i < bassedata.ilist.Count; i++)
                {
                    if (tempstr == bassedata.ilist[i].pn)
                    {
                        wsWeekly.Cells[wsWeeklyLine, 4] = bassedata.ilist[i].type;
                        wsWeekly.Cells[wsWeeklyLine, 6] = bassedata.ilist[i].name;
                        wsWeekly.Cells[wsWeeklyLine, 8] = bassedata.ilist[i].discription;
                        break;
                    }
                }*/
                if (bassedata.ilist.ContainsKey(tempstr))
                {
                    wsWeekly.Cells[wsWeeklyLine, 4] = bassedata.ilist[tempstr].type;
                    wsWeekly.Cells[wsWeeklyLine, 6] = bassedata.ilist[tempstr].name;
                    wsWeekly.Cells[wsWeeklyLine, 8] = bassedata.ilist[tempstr].discription;
                }

                wsWeekly.Cells[wsWeeklyLine, 5] = "Wistron";
                wsWeekly.Cells[wsWeeklyLine, 28] = "Y";
                wsWeekly.Cells[wsWeeklyLine, 29] = "Lily";
                wsWeekly.Cells[wsWeeklyLine, 30] = "HGY";
                wsWeekly.Cells[wsWeeklyLine, 33] = "Jimmy";
                wsWeekly.Cells[wsWeeklyLine, 34] = "Finished";

            }
            if (custermorID == 5)//LCFC
            {
                if (bassedata.DailyReportDate_Weekly != -1)
                {
                    wsWeekly.Cells[wsWeeklyLine, bassedata.DailyReportDate_Weekly] = DailyReportDate.Year.ToString() + "-" + DailyReportDate.Month.ToString("D2") + "-" + DailyReportDate.Day.ToString("D2");
                }
                wsWeekly.Cells[wsWeeklyLine, bassedata.Decision_Weekly] = wsDaily.Cells[wsDailyLine, bassedata.Decision_Daily].Text;
                wsWeekly.Cells[wsWeeklyLine, 4] = wsDaily.Cells[wsDailyLine, 4].Text;
                wsWeekly.Cells[wsWeeklyLine, 5] = "FCT";
            }
            return true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //button4.Enabled = true;
            //button5.Enabled = true;
            button6.Enabled = true;

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            button6.Enabled = true;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            button6.Enabled = true;
        }
        

    }
    
    public class Identifier
    {
        public int lineno = 1;
        public string sn = "";
        public int year = 0;
        public int month = 0;
        public int day = 0;
    }
    public class BaseData
    {
        //public List<ADDDATA> ilist = new List<ADDDATA>();
        public Dictionary<string,ADDDATA> ilist = new Dictionary<string, ADDDATA>();
        int LastCustomID;

        public class ADDDATA
        {
            public string pn = "";
            public string type = "";
            public string name = "";
            public string discription = "";
            public string modeName = "";
        }
        public BaseData()
        {
            LastCustomID = -1;
        }
        private void InitAdditionalData(int type)
        {
            string tempstr;
            ilist.Clear(); //c#中，不再像c++一样手动delete原来new的对象，c#垃圾回收机制自动释放内存
            //type: 0 msi; 1:usi/ 2:wistron/ 3:wistron aio/ 4:inventec/ 5:lcfc
            if (type == 1)
            {
                Excel.Application excel = null;
                Excel.Workbooks wbooks = null;
                Excel.Workbook wbook = null;
                Excel.Worksheet wsheet = null;
                object Nothing = System.Reflection.Missing.Value;
                try
                {

                    excel = new Excel.Application();
                    excel.UserControl = true;
                    excel.DisplayAlerts = false;

                    wbooks = excel.Workbooks;
                    wbook = wbooks.Open("E:\\work\\TSC\\Report\\Foundamental data\\PN list USI Lenovo Server Module list_UPDATE.xlsx", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
                    wsheet = wbook.Sheets[1];

                    for (int i = 2; i <= wsheet.UsedRange.Rows.Count; i++)
                    {
                        if (i == 334)
                            tempstr = "";
                        tempstr = ((Excel.Range)wsheet.Cells[i, 3]).Text;
                        tempstr = tempstr.Trim();
                        if (tempstr == "")
                        {
                            break;
                        }
                        ADDDATA adddata = new ADDDATA();
                        adddata.pn = tempstr;

                        tempstr = ((Excel.Range)wsheet.Cells[i, 4]).Text;
                        tempstr = tempstr.Trim();
                        adddata.type = tempstr;

                        tempstr = ((Excel.Range)wsheet.Cells[i, 5]).Text;
                        tempstr = tempstr.Trim();
                        adddata.name = tempstr;

                        tempstr = ((Excel.Range)wsheet.Cells[i, 12]).Text;
                        tempstr = tempstr.Trim();
                        adddata.modeName = tempstr;

                        //ilist.Add(adddata.pn,adddata);
                        ilist[adddata.pn] = adddata;
                    }

                }
                catch (Exception mye)
                {

                    MessageBox.Show(mye.ToString() + "  open " + "PN List USI file failed");
                    return;
                }
                finally
                {
                    if (excel != null)
                    {
                        if (wbooks != null)
                        {
                            if (wbook != null)
                            {
                                if (wsheet != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsheet);
                                    wsheet = null;
                                }
                                wbook.Close(false, Nothing, Nothing);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbook);
                                wbook = null;
                            }
                            wbooks.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wbooks);
                            wbooks = null;
                        }
                        excel.Application.Workbooks.Close();
                        excel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                        excel = null;
                    }

                }

            }


            if (type == 2)
            {
                Excel.Application excel = null;
                Excel.Workbooks wbooks = null;
                Excel.Workbook wbook = null;
                Excel.Worksheet wsheet = null;
                object Nothing = System.Reflection.Missing.Value;
                try
                {
                    
                    excel = new Excel.Application();
                    excel.UserControl = true;
                    excel.DisplayAlerts = false;

                    wbooks = excel.Workbooks;
                    wbook = wbooks.Open("E:\\work\\TSC\\Report\\Foundamental data\\PN List Wistron 09-Feb-2022.xlsx", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
                    wsheet = wbook.Sheets[1];

                    for (int i = 2; i <= wsheet.UsedRange.Rows.Count; i++)
                    {
                        if (i == 334)
                            tempstr = "";
                        tempstr = ((Excel.Range)wsheet.Cells[i, 1]).Text;
                        tempstr = tempstr.Trim();
                        if (tempstr == "")
                        {
                            break;
                        }
                        ADDDATA adddata = new ADDDATA();
                        adddata.pn = tempstr;

                        tempstr = ((Excel.Range)wsheet.Cells[i, 4]).Text;
                        tempstr = tempstr.Trim();
                        adddata.type = tempstr;

                        tempstr = ((Excel.Range)wsheet.Cells[i, 2]).Text;
                        tempstr = tempstr.Trim();
                        adddata.name = tempstr;
                        tempstr = ((Excel.Range)wsheet.Cells[i, 3]).Text;
                        tempstr = tempstr.Trim();
                        adddata.discription = tempstr;
                        if (tempstr == "MB")
                        {
                            if (adddata.type == "")
                                adddata.type = "ECAT planar";
                        }
                        else if (tempstr == "Card")
                        {
                            if (adddata.type == "")
                                adddata.type = "ECAT card";
                        }
                        //ilist.Add(adddata.pn,adddata);
                        ilist[adddata.pn] = adddata;
                    }

                }
                catch (Exception mye)
                {

                    MessageBox.Show(mye.ToString() + "  open " + "PN List Wistron file failed");
                    return;
                }
                finally
                {
                    if (excel != null)
                    {
                        if (wbooks != null)
                        {
                            if (wbook != null)
                            {
                                if (wsheet != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsheet);
                                    wsheet = null;
                                }
                                wbook.Close(false, Nothing, Nothing);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbook);
                                wbook = null;
                            }
                            wbooks.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wbooks);
                            wbooks = null;
                        }
                        excel.Application.Workbooks.Close();
                        excel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                        excel = null;
                    }

                }

            }

            if (type == 4)  //Inventec
            {
                Excel.Application excel = null;
                Excel.Workbooks wbooks = null;
                Excel.Workbook wbook = null;
                Excel.Worksheet wsheet = null;
                object Nothing = System.Reflection.Missing.Value;
                try
                {

                    excel = new Excel.Application();
                    excel.UserControl = true;
                    excel.DisplayAlerts = false;

                    wbooks = excel.Workbooks;
                    wbook = wbooks.Open("E:\\work\\TSC\\Report\\Foundamental data\\PN list inventec.xlsx", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
                    wsheet = wbook.Sheets[1];

                    for (int i = 2; i <= wsheet.UsedRange.Rows.Count; i++)
                    {
                        tempstr = ((Excel.Range)wsheet.Cells[i, 2]).Text;
                        tempstr = tempstr.Trim();
                        if (tempstr == "")
                            break;
                        ADDDATA adddata = new ADDDATA();
                        adddata.pn = tempstr;
                        tempstr = ((Excel.Range)wsheet.Cells[i, 1]).Text;
                        tempstr = tempstr.Trim();
                        adddata.name = tempstr;

                        tempstr = ((Excel.Range)wsheet.Cells[i, 3]).Text;
                        tempstr = tempstr.Trim();
                        adddata.type = tempstr;
                        //ilist.Add(adddata.pn, adddata);
                        ilist[adddata.pn] = adddata;
                    }

                }
                catch (Exception mye)
                {

                    MessageBox.Show(mye.ToString() + "  open " + "PN List Inventec file failed");
                    return;
                }
                finally
                {
                    if (excel != null)
                    {
                        if (wbooks != null)
                        {
                            if (wbook != null)
                            {
                                if (wsheet != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsheet);
                                    wsheet = null;
                                }
                                wbook.Close(false, Nothing, Nothing);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbook);
                                wbook = null;
                            }
                            wbooks.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wbooks);
                            wbooks = null;
                        }
                        excel.Application.Workbooks.Close();
                        excel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                        excel = null;
                    }

                }

            }
        }
        public bool Init(int type)
        {
            //type: 0 msi; 1:usi/ 2:wistron/ 3:wistron aio/ 4:inventec/ 5:lcfc
            if (LastCustomID == type)
                return true;

            InitAdditionalData(type);

            if (type == 0) //MSI
            {
                SheetID_Weekly = 2;

                PN_Daily = 3;
                SN_Daily = 5;
                FailDate_Daily = 7;
                Failure_Daily = 6;
                FaResult_Daily = 8;
                Decision_Daily = 10;

                DailyReportDate_Weekly = 2;
                PN_Weekly = 8;
                SN_Weekly = 9;//and 16
                FailDate_Weekly = 11;
                Failure_Weekly = 13;
                FaResult_Weekly = 16;
                Decision_Weekly = 23;

                Disposition_Weekly = -1;
                RtvBitmap_Weekly = -1;
                CidBitmap_Weekly = -1;
                NdfBitmap_Weekly = -1;
                FwfBitmap_Weekly = -1;
                Responsibility_Weekly = -1;

                Weeknum_Weekly = -1;
            }
            if (type == 1) //USI
            {
                SheetID_Weekly = 1;

                PN_Daily = 3;
                SN_Daily = 5;
                FailDate_Daily = 7;
                Failure_Daily = 6;
                FaResult_Daily = 8;
                Decision_Daily = 10;

/*                //insert 2 columns after 6th column, so every column>7 need to add 2
                DailyReportDate_Weekly = 4;
                PN_Weekly = 7;
                SN_Weekly = 8 + 2;
                FailDate_Weekly = 3;
                Failure_Weekly = 9 + 2;
                FaResult_Weekly = 11 + 2;
                Decision_Weekly = 12 + 2;

                Disposition_Weekly = 19 + 2;
                RtvBitmap_Weekly = 13 + 2;
                CidBitmap_Weekly = 14 + 2;
                NdfBitmap_Weekly = 15 + 2;
                FwfBitmap_Weekly = 16 + 2;
                Responsibility_Weekly = -1;
*/
                //insert 2 columns after 6th column, so every column>7 need to add 2
                DailyReportDate_Weekly = -1;
                PN_Weekly = 5;
                SN_Weekly = 8;
                FailDate_Weekly = 3;
                Failure_Weekly = 9;
                FaResult_Weekly = 10;
                Decision_Weekly = 11;

                Disposition_Weekly = -1;
                RtvBitmap_Weekly = -1;
                CidBitmap_Weekly = -1;
                NdfBitmap_Weekly = -1;
                FwfBitmap_Weekly = -1;
                Responsibility_Weekly = -1;

                Weeknum_Weekly = 2;
            }

            if (type == 2) //wistron
            {
                SheetID_Weekly = 1;

                PN_Daily = 3;
                SN_Daily = 5;
                FailDate_Daily = 7;
                FaResult_Daily = 8;
                Failure_Daily = 6;
                Decision_Daily = 10;

                DailyReportDate_Weekly = 3;
                SN_Weekly = 9;//and 16
                PN_Weekly = 7;
                FailDate_Weekly = 11;
                FaResult_Weekly = 17;
                Failure_Weekly = 12;
                Decision_Weekly = 19;

                Disposition_Weekly = 20;  
                RtvBitmap_Weekly = 22;  
                CidBitmap_Weekly = 23;
                NdfBitmap_Weekly = 24;
                FwfBitmap_Weekly = 25;
                Responsibility_Weekly = 31;

                Weeknum_Weekly = 2;
            }

            if (type == 4) //Inventec
            {
                SheetID_Weekly = 2;

                PN_Daily = 7;
                SN_Daily = 8;
                FailDate_Daily = 9;
                Failure_Daily = 10;
                FaResult_Daily = 15;
                Decision_Daily = 16;

                DailyReportDate_Weekly = 3;
                PN_Weekly = 7;
                SN_Weekly = 8;//and 16
                FailDate_Weekly = 9;
                Failure_Weekly = 10;
                FaResult_Weekly = 15;
                Decision_Weekly = 16;

                Disposition_Weekly = -1;
                RtvBitmap_Weekly = -1;
                CidBitmap_Weekly = -1;
                NdfBitmap_Weekly = -1;
                FwfBitmap_Weekly = -1;
                Responsibility_Weekly = -1;

                Weeknum_Weekly = 2;
            }
            if (type == 5) //LCFC
            {
                SheetID_Weekly = 1;

                PN_Daily = 3;
                SN_Daily = 5;
                FailDate_Daily = 7;
                Failure_Daily = 6;
                FaResult_Daily = 9;
                Decision_Daily = 11;

                DailyReportDate_Weekly = -1;
                PN_Weekly = 3;
                SN_Weekly = 2;
                FailDate_Weekly = 8;
                Failure_Weekly = 6;
                FaResult_Weekly = 7;
                Decision_Weekly = 9;

                Disposition_Weekly = -1;
                RtvBitmap_Weekly = -1;
                CidBitmap_Weekly = -1;
                NdfBitmap_Weekly = -1;
                FwfBitmap_Weekly = -1;
                Responsibility_Weekly = -1;

                Weeknum_Weekly = 1;
            }
            LastCustomID = type;
            return true;

        }

        //weekly report data sheet index
        private int sheetid_weekly;
        public int SheetID_Weekly
        {
            get { return sheetid_weekly; }
            set { sheetid_weekly = value; }
        }

        //SN
        private int sn_weekly;
        public int SN_Weekly
        {
            get { return sn_weekly; }
            set { sn_weekly = value; }
        }
        private int sn_daily;
        public int SN_Daily
        {
            get { return sn_daily; }
            set { sn_daily = value; }
        }
        //PN
        private int pn_weekly;
        public int PN_Weekly
        {
            get { return pn_weekly; }
            set { pn_weekly = value; }
        }
        private int pn_daily;
        public int PN_Daily
        {
            get { return pn_daily; }
            set { pn_daily = value; }
        }
        //Fail Date
        private int faildate_weekly;
        public int FailDate_Weekly
        {
            get { return faildate_weekly; }
            set { faildate_weekly = value; }
        }
        private int faildate_daily;
        public int FailDate_Daily
        {
            get { return faildate_daily; }
            set { faildate_daily = value; }
        }
        //Fa result
        private int faresult_weekly;
        public int FaResult_Weekly
        {
            get { return faresult_weekly; }
            set { faresult_weekly = value; }
        }
        private int faresult_daily;
        public int FaResult_Daily
        {
            get { return faresult_daily; }
            set { faresult_daily = value; }
        }
        //Failure
        private int failure_weekly;
        public int Failure_Weekly
        {
            get { return failure_weekly; }
            set { failure_weekly = value; }
        }
        private int Failure_daily;
        public int Failure_Daily
        {
            get { return Failure_daily; }
            set { Failure_daily = value; }
        }
        //Decision
        private int decision_weekly;
        public int Decision_Weekly
        {
            get { return decision_weekly; }
            set { decision_weekly = value; }
        }
        private int decision_daily;
        public int Decision_Daily
        {
            get { return decision_daily; }
            set { decision_daily = value; }
        }

        private int disposition_weekly;
        public int Disposition_Weekly
        {
            get { return disposition_weekly; }
            set { disposition_weekly = value; }
        }
        private int rtvbitmap_weekly;
        public int RtvBitmap_Weekly
        {
            get { return rtvbitmap_weekly; }
            set { rtvbitmap_weekly = value; }
        }
        private int cidbitmap_weekly;
        public int CidBitmap_Weekly
        {
            get { return cidbitmap_weekly; }
            set { cidbitmap_weekly = value; }
        }
        private int ndfbitmap_weekly;
        public int NdfBitmap_Weekly
        {
            get { return ndfbitmap_weekly; }
            set { ndfbitmap_weekly = value; }
        }
        private int fwbitmap_weekly;
        public int FwfBitmap_Weekly
        {
            get { return fwbitmap_weekly; }
            set { fwbitmap_weekly = value; }
        }
        private int responsibility_weekly;
        public int Responsibility_Weekly
        {
            get { return responsibility_weekly; }
            set { responsibility_weekly = value; }
        }
        private int weeknum_weekly;
        public int Weeknum_Weekly
        {
            get { return weeknum_weekly; }
            set { weeknum_weekly = value; }
        }
        private int dailyreportdate_weekly;
        public int DailyReportDate_Weekly
        {
            get { return dailyreportdate_weekly; }
            set { dailyreportdate_weekly = value; }
        }

    }
}
