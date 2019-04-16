using System;
using System.Configuration;
using System.Runtime.InteropServices;
using System.IO;
using System.Net.Mail;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;




namespace Klarna
{
    class Program
    {


        static Excel.Application xlApp;
        static Excel.Workbook xlWorkBook;
        static Excel.Worksheet xlWorkSheet;
        static Excel.Range range;

        static int columnCnt;
        static int rowCnt;
        static int output;
        static int output1;

        private static readonly string CsvFileName = ConfigurationManager.AppSettings["csvFileName"];
        private static readonly string TsvFileName = ConfigurationManager.AppSettings["tsvFileName"];
        private static readonly string ExcelFileName = ConfigurationManager.AppSettings["excelFileName"];

        static void Main(string[] args)
        {
            

            FileInfo fi1 = new FileInfo(CsvFileName);
            Delete(TsvFileName);
            ConvertCSVtoTabDelimited(fi1);



            FileInfo fi2 = new FileInfo(TsvFileName);
            Delete(ExcelFileName);
            ConvertTSVtoEXCEL(fi2);




            if (File.Exists(ExcelFileName))
            {
                string filenam = ConfigurationManager.AppSettings["MyPath1"] + DateTime.Now.ToString("dd_MM_yyyy") + ".xlsx";

                Delete(filenam);

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(ExcelFileName);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                columnCnt = range.Columns.Count;
                rowCnt = range.Rows.Count;



                int IndexOfCol1 = DeleteColumn("TargetTeam");
                int IndexOfCol2 = DeleteColumn("batchNumber");



                if (IndexOfCol1 > 0 && IndexOfCol2 > 0)
                {
                    ((Excel.Range)xlWorkSheet.Columns[IndexOfCol1]).EntireColumn.Delete(null);

                    ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Columns[IndexOfCol2 - 1]).EntireColumn.Delete(null);

                    int FormatCol1 = FormatColumn("InvoiceNumber");
                    int FormatCol2 = FormatColumn("ReceiptId");
                    int FormatCol3 = FormatColumn("dateentered");
                    int FormatCol4 = FormatColumn("VoidHeaderId");






                    if (FormatCol1 > 0 && FormatCol2 > 0 && FormatCol3 > 0 && FormatCol4 > 0)
                    {

                        xlWorkSheet.Columns[FormatCol1].NumberFormat = Constants.Format;
                        xlWorkSheet.Columns[FormatCol2].NumberFormat = Constants.Format;
                        xlWorkSheet.Columns[FormatCol3].NumberFormat = Constants.FormatDate;
                        xlWorkSheet.Columns[FormatCol4].NumberFormat = Constants.Format;



                      //  Display();

                        string filename = ConfigurationManager.AppSettings["MyPath1"] + DateTime.Now.ToString("dd_MM_yyyy");
                        xlWorkBook.SaveAs(filename + ".xlsx");

                       



                        Marshal.ReleaseComObject(range);

                        Marshal.ReleaseComObject(xlWorkSheet);//close and release

                        xlWorkBook.Close();

                        Marshal.ReleaseComObject(xlWorkBook);//quit and release

                        xlApp.Quit();

                        Marshal.ReleaseComObject(xlApp);

                        sendemailAsync("sagar.andre@asos.com", filename + ".xlsx");
                    }
                    else
                    {

                        Console.WriteLine(Constants.FormattingColumnNotExists);
                    }
                }
                else
                {

                    Console.WriteLine(Constants.DeletingColumnNotexists);

                }
            }
            else
            {

                Console.WriteLine(Constants.FileNotExists);

            }

            Console.ReadKey();
        }
        private static void ConvertCSVtoTabDelimited(FileInfo fi)// csv-tsv
        {
            try
            {
                string NewFileName = Path.Combine(Path.GetDirectoryName(fi.FullName), Path.GetFileNameWithoutExtension(fi.FullName) + ".tsv");
                System.IO.File.WriteAllText(NewFileName, System.IO.File.ReadAllText(fi.FullName).Replace(",", "\t"));
            }
            catch (Exception ex)
            {
                Console.WriteLine("File: " + fi.FullName + Environment.NewLine + Environment.NewLine + ex.ToString(), "Error Converting File");
            }
        }
        private static void ConvertTSVtoEXCEL(FileInfo fi1)
        {
            string worksheetsName = "sheet1";


            var format = new ExcelTextFormat();
            format.Delimiter = '\t';
            format.EOL = "\r";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(ExcelFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                worksheet.Cells["A1"].LoadFromText(new FileInfo(TsvFileName), format);
                package.Save();
            }


        }

        private static void Delete(string fn)
        {
            if (File.Exists(fn))
            {
                File.Delete(fn);
            }
        }

        private static void Display()
        {
            for (int row = 1; row <= rowCnt; row++)

            {
                for (int col = 1; col <= columnCnt; col++)
                {
                    if (col == 1)

                        Console.Write("\r\n");

                    if (range.Cells[row, col] != null && range.Cells[row, col].Value2 != null)

                        Console.Write(range.Cells[row, col].Value2.ToString() + "\t");
                }
            }

        }



        private static int DeleteColumn(string str)
        {

            for (int col = 1; col <= columnCnt; col++)
            {

                if (xlWorkSheet.Cells[1, col].value == str)
                {
                    output1 = col;
                    return output1;
                }

            }

            return 0;
        }
        private static int FormatColumn(string str)
        {

            for (int col = 1; col <= columnCnt; col++)

            {
                if (xlWorkSheet.Cells[1, col].value == str)
                {
                    output = col;
                    return output;
                }
            }

            return 0;
        }


        private static void sendemailAsync(string Toemail, string filePath)
        {


          

            string subject = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " Daily Klarna Refunds Report for Klarna (automated) (" + rowCnt +
                         " Results Found )  ";



            var SmtpServer = new SmtpClient("smtp.office365.com");
           
            var froMailAddress = new MailAddress("sagar.andre@asos.com");
            var toMailAddress = new MailAddress(Toemail);

            var mailMessage = new MailMessage(froMailAddress, toMailAddress)
            {
                Subject = subject,
               
            };
            Attachment attachment;
            attachment = new Attachment(filePath);
            mailMessage.Attachments.Add(attachment);
            mailMessage.CC.Add("keyur.patel@asos.com");
            mailMessage.CC.Add("yogesh.jadhav@asos.com");
            mailMessage.Body = @"Hi Klarna Support, <br /><br />In the attached spreadsheet please find today's list of possibly failed refunds.<br />As usual please could you review each of these and let us know if they are at a failed or success status so we can action accordingly.<br />Any refunds which have failed we will retry on our side and any refunds which are successful we will set to complete within Back Office.<br />Please endeavor to reply back to this email within 24 hours so we can action accordingly.<br /><br /><br />Regards,<br />Sagar Andre";
            mailMessage.IsBodyHtml = true;
          
            SmtpServer.EnableSsl = true;
            
            SmtpServer.Send(mailMessage);

        }


    }

}











