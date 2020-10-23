using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SBPD_DT_Sentinel
{
    public partial class Form1 : Form
    {
        List<DataTicketCitation> dataticketCites = new List<DataTicketCitation>();

        List<DataTicketCitation> oneHourCites = new List<DataTicketCitation>();
        List<DataTicketCitation> sweeperCites = new List<DataTicketCitation>();
        List<DataTicketCitation> mainStCites = new List<DataTicketCitation>();
        List<DataTicketCitation> beachLotCites = new List<DataTicketCitation>();
        List<DataTicketCitation> otherCites = new List<DataTicketCitation>();
        public Form1()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            InitializeComponent();
        }

        private void importDTData(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];
                var start = firstWorksheet.Dimension.Start;
                var end = firstWorksheet.Dimension.End;
                for (int row = start.Row + 1; row <= end.Row - 1; row++)
                {
                    string citationDateTime = firstWorksheet.Cells[row, 1].Text;
                    string citationNumber = firstWorksheet.Cells[row, 2].Text;
                    string citationViolationCode = firstWorksheet.Cells[row, 3].Text;
                    string citationLocation = firstWorksheet.Cells[row, 4].Text;
                    string citationComment = firstWorksheet.Cells[row, 6].Text;
                    dataticketCites.Add(new DataTicketCitation(citationDateTime, citationNumber, citationViolationCode, citationLocation, citationComment));
                }
            }

            foreach (DataTicketCitation citation in dataticketCites)
            {
                if (citation.violationCode.Contains("8.15.055 SBMC") && citation.comments.Contains("ONE HOUR"))
                {
                    oneHourCites.Add(citation);
                }
                else if (citation.violationCode.Contains("8.15.055 SBMC") && citation.comments.Contains("2 HRS"))
                {
                    mainStCites.Add(citation);
                }
                else if (citation.location.Contains("MAIN ST"))
                {
                    mainStCites.Add(citation);
                } else if (citation.comments.Contains("STREET SWEEPING"))
                {
                    sweeperCites.Add(citation);
                } else if (citation.location.Contains("BEACH LOT"))
                {
                    beachLotCites.Add(citation);
                } else
                {
                    otherCites.Add(citation);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialogDT.FileName = "download"; // Default file name
            openFileDialogDT.DefaultExt = ".xls"; // Default file extension
            openFileDialogDT.Filter = "Excel files (.xlsx)|*.xlsx"; // Filter files by extension
            openFileDialogDT.ShowDialog();
            label1.Visible = true;
            string fileName = openFileDialogDT.FileName;
            importDTData(fileName);
            /*
            using (WebClient wc = new WebClient())
            {
                string location = dataticketCites[1].location;
                var json = wc.DownloadString("http://dev.virtualearth.net/REST/V1/Routes/Walking?wp.0=" + location + "&wp.1=" + "8TH ST BEACH LOT" + "&key=AoyuBnIo4_lBOPShKB--TkfzosE0nGDN1gkJX9sBl5XkA3Tz7bC0xzofuSCfD7PN");
                dynamic jsonObj = JsonConvert.DeserializeObject(json);
                richTextBox1.Text = jsonObj.resourceSets[0].resources[0].travelDistance;
            }
            */

        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialogDT.FileName = "Monthly Log"; // Default file name
            openFileDialogDT.DefaultExt = ".xls"; // Default file extension
            openFileDialogDT.Filter = "Excel files (.xlsx)|*.xlsx"; // Filter files by extension
            openFileDialogDT.ShowDialog();
            label2.Visible = true;
            string fileName = openFileDialogDT.FileName;

            FileInfo fi = new FileInfo(fileName);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];
                var start = firstWorksheet.Dimension.Start;
                var end = firstWorksheet.Dimension.End;
                if (firstWorksheet.Name == "CSO MONTHLY")
                {
                    for (int row = start.Row + 6; row <= 37; row++)
                    {
                        int citeNumber = 0;
                        // ONE HOUR CITES
                        foreach (DataTicketCitation oneHourCite in oneHourCites)
                        {
                            string[] convDateTime = oneHourCite.dateTime.Split("/");
                            string day = convDateTime[1];
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day.Length == 1)
                                day = "0" + day;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 12].Value = citeNumber;
                        citeNumber = 0;

                        // SWEEPER CITES
                        foreach (DataTicketCitation sweeperCite in sweeperCites)
                        {
                            string[] convDateTime = sweeperCite.dateTime.Split("/");
                            string day = convDateTime[1];
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day.Length == 1)
                                day = "0" + day;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 13].Value = citeNumber;
                        citeNumber = 0;

                        // MAIN ST CITES
                        foreach (DataTicketCitation mainStCite in mainStCites)
                        {
                            string[] convDateTime = mainStCite.dateTime.Split("/");
                            string day = convDateTime[1];
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day.Length == 1)
                                day = "0" + day;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 14].Value = citeNumber;
                        citeNumber = 0;

                        // BEACH LOT CITES
                        foreach (DataTicketCitation beachLotCite in beachLotCites)
                        {
                            string[] convDateTime = beachLotCite.dateTime.Split("/");
                            string day = convDateTime[1];
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day.Length == 1)
                                day = "0" + day;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 15].Value = citeNumber;
                        citeNumber = 0;


                        // OTHER CITES
                        foreach (DataTicketCitation otherCite in otherCites)
                        {
                            string[] convDateTime = otherCite.dateTime.Split("/");
                            string day = convDateTime[1];
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day.Length == 1)
                                day = "0" + day;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 16].Value = citeNumber;
                        citeNumber = 0;
                    }
                    excelPackage.Save();
                }
            }
        }
    }
}
