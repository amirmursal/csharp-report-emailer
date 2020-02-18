using System;
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
            this.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            this.btn_upload_excel.Enabled = false;
        }

        string doctorColumn = "";

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
                                //DirectoryClean();
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
                                //DirectoryClean();
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
                //string[] FilterList = new string[] { "Dr. Anuj Camanocha", "Damon M Barbieri" , "DR Baptiste", "Dr Kerry White Brown", "Dr Shafeena Chatur", "Dr. Dunn", "Dr. Grob", "Dr. McGill", "Dr. Williams", "Dr.Bahr", "Dr.chawla", "GSO", "Hickory Tops" };
                string[] FilterList = new string[] { "Dr. Anuj Camanocha" };

                foreach (string pdf in FilterList)
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

                    string remark = "Updated";

                    int docCol = Int32.Parse(doctorColumn);
                    excelRange.AutoFilter(docCol, pdf,
                                     Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                    excelRange.AutoFilter(16,remark,
                                     Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);



                    string excelFileDestination = Path.Combine(Directory.GetCurrentDirectory(), "ExcelFiles");

                    CreateExcel(pdf);

                    string destPath = excelFileDestination + "\\" + pdf + ".xlsx";
                    Workbook destworkBook = _excelApp.Workbooks.Open(destPath, 0, false);
                    Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);

                    destworkSheet.Rows.RowHeight = 25;

                    destworkSheet.Columns.ColumnWidth = 25;

                    destworkSheet.Columns[13].ColumnWidth = 75;

                    Range from = worksheet.UsedRange;             
                    Range to = destworkSheet.UsedRange;               

                    from.Copy(to);
                    _excelApp.DisplayAlerts = false;
                    destworkBook.SaveAs(excelFileDestination + "\\" + pdf + ".xlsx");

                    string attachmentFile = excelFileDestination + "\\" + pdf + ".xlsx";


                    destworkBook.Close(false, Type.Missing, Type.Missing);


                    SendEmail(attachmentFile);

                    workbook.Close(false, Type.Missing, Type.Missing);

                    _excelApp.Quit();
                }
                getNonUpdatedData(file);

                MessageBox.Show("EXcel Generated Successfully");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void getNonUpdatedData(string file)
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

            excelRange.Worksheet.AutoFilterMode = false;

            string remark = "Non-Updated";

            string[] listFilter = new string[] { "<>Updated*" };
                      

            excelRange.AutoFilter(16, listFilter,
                             Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

            string excelFileDestination = Path.Combine(Directory.GetCurrentDirectory(), "ExcelFiles");

            CreateExcel(remark);

            string destPath = excelFileDestination + "\\" + remark + ".xlsx";
            Workbook destworkBook = _excelApp.Workbooks.Open(destPath, 0, false);
            Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);

            destworkSheet.Rows.RowHeight = 25;

            destworkSheet.Columns.ColumnWidth = 25;

            destworkSheet.Columns[13].ColumnWidth = 75;

            Range from = worksheet.UsedRange;
            Range to = destworkSheet.UsedRange;

            from.Copy(to);
            _excelApp.DisplayAlerts = false;
            destworkBook.SaveAs(excelFileDestination + "\\" + remark + ".xlsx");

            string attachmentFile = excelFileDestination + "\\" + remark + ".xlsx";


            destworkBook.Close(false, Type.Missing, Type.Missing);          

            workbook.Close(false, Type.Missing, Type.Missing);

            _excelApp.Quit();
        }

        public void CreateExcel(string pdf)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


            //Here saving the file in xlsx
            string excelFileDestination = Path.Combine(Directory.GetCurrentDirectory(), "ExcelFiles");
            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(excelFileDestination + "\\" + pdf + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }

        public void SendEmail(string attachementFile)
        {
            try
            {
                var htmlBody = "<p>Please find Attached document</p>";
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress("amirmursal@gmail.com");
                message.To.Add(new MailAddress("amirthink72@gmail.com"));
                message.Subject = "Test";
                message.IsBodyHtml = true; //to make message body as html  
                message.Body = htmlBody;
                message.Attachments.Add(new Attachment(attachementFile));
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = comboBox1.SelectedIndex;
            Object selectedItem = comboBox1.SelectedItem;
            doctorColumn = selectedItem.ToString();
            this.btn_upload_excel.Enabled = true;
        }
    }

}
