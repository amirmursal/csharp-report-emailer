using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Net;

namespace MediNetCorpPDF_Extractor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DirectoryClean();
        }

        public void DirectoryClean()
        {
            try
            {
                string ExcelFiles = Path.Combine(Directory.GetCurrentDirectory() + @"\ExcelFiles");
                string[] AllExcelFiles = Directory.GetFiles(ExcelFiles);
                foreach (string file in AllExcelFiles)
                {
                    File.Delete(file);
                }               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void btn_upload_excel_Click(object sender, EventArgs e)
        {
            try
            {
                string excelFileDestination = Path.Combine(Directory.GetCurrentDirectory() + @"\ExcelFiles");
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Multiselect = true;
                if (Directory.Exists(excelFileDestination))
                {
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        foreach (string file in dlg.FileNames)
                        {
                            if (file.Contains(".xlsx") || file.Contains(".xls"))
                            {
                                File.Copy(file, Path.Combine(excelFileDestination, Path.GetFileName(file)), true);
                                MessageBox.Show(file + " uploaded  Successfully");
                                ProcessExcel(file);
                                DirectoryClean();
                            }

                            else
                            {
                                MessageBox.Show("Upload Excel files only");
                                return;
                            }
                        }
                    }
                }
                else
                {
                    System.IO.Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\ExcelFiles");
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        foreach (string file in dlg.FileNames)
                        {
                            if (file.Contains(".xlsx") || file.Contains(".xls"))
                            {
                                File.Copy(file, Path.Combine(excelFileDestination, Path.GetFileName(file)), true);
                                MessageBox.Show(file + " uploaded  Successfully");
                                ProcessExcel(file);
                                DirectoryClean();
                            }
                            else
                            {
                                MessageBox.Show("Upload Excel files only");
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public void ProcessExcel(string file)
        {

            try
            {  
                //create the Application object we can use in the member functions.
                Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
                _excelApp.Visible = true;

                //open the workbook
                Workbook workbook = _excelApp.Workbooks.Open(file,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                //find the used range in worksheet
                Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);


                var cellValue = (worksheet.Cells[2, 2] as Microsoft.Office.Interop.Excel.Range).Text;

                int rowCount = worksheet.UsedRange.Rows.Count;
                int columnCount = worksheet.UsedRange.Columns.Count;

                List<string> columnValue = new List<string>();
                List<string> BookmarkColumnValue = new List<string>();
                Microsoft.Office.Interop.Excel.Range visibleCells = worksheet.UsedRange.SpecialCells(
                                 Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible,
                           Type.Missing);
                var dictionary = new List<KeyValuePair<string, string>>();
                var providerDictionary = new List<KeyValuePair<string, string>>();
                List<Tuple<string, string, string>> list = new List<Tuple<string, string, string>>();
                foreach (Microsoft.Office.Interop.Excel.Range area in visibleCells.Areas)
                {
                    foreach (Microsoft.Office.Interop.Excel.Range row in area.Rows)
                    {
                        string softwareName = ((Microsoft.Office.Interop.Excel.Range)row.Cells[2, 2]).Text;
                        if (softwareName != ""){
                            
                            SendEmail(softwareName);
                        }
                    }
                }
                MessageBox.Show("All Report has been send to perticular doctors");
                workbook.Close(false, Type.Missing, Type.Missing);
                _excelApp.Quit();
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void SendEmail(string softwareName)
        {
            try
            {
                var htmlBody = "<h1>"+ softwareName + "</h1>";
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress("amirmursal@gmail.com");
                message.To.Add(new MailAddress("amirthink72@gmail.com"));
                message.Subject = "Test";
                message.IsBodyHtml = true; //to make message body as html  
                message.Body = htmlBody;              
                smtp.Host = "smtp.gmail.com"; //for gmail host  
                smtp.Port = 587;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("amirmursal@gmail.com", "amirarshin");
                smtp.EnableSsl = true;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }       

    }

}
