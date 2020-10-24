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
        List<DataTicketCitation> otherCites = new List<DataTicketCitation>();

        List<DataTicketCitation> sweeperCites = new List<DataTicketCitation>();
        List<DataTicketCitation> mainStCites = new List<DataTicketCitation>();
        List<DataTicketCitation> beachLotCites = new List<DataTicketCitation>();

        List<DataTicketCitation> oldTownCites = new List<DataTicketCitation>();
        List<DataTicketCitation> northEndCites = new List<DataTicketCitation>();
        List<DataTicketCitation> theHillCites = new List<DataTicketCitation>();
        List<DataTicketCitation> handicapCites = new List<DataTicketCitation>();
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
                    oldTownCites.Add(citation);
                }
                else if (citation.violationCode.Contains("8.15.055 SBMC") && citation.comments.Contains("2 HRS"))
                {
                    mainStCites.Add(citation);
                }
                else if (citation.violationCode.Contains("8.15.055 SBMC") && citation.comments.Contains("VEHICLE FAILED TO MOVE THE REQUIRED 150 FEET"))
                {
                    mainStCites.Add(citation);
                }
                else if (citation.location.Contains("MAIN ST") || citation.location.Contains("METERED ZONE"))
                {
                    mainStCites.Add(citation);
                }
                else if (citation.comments.Contains("STREET SWEEPING"))
                {
                    sweeperCites.Add(citation);
                }
                else if (citation.location.Contains("BEACH LOT"))
                {
                    beachLotCites.Add(citation);
                }
                else if (citation.violationCode.Contains("22507.8") || citation.violationCode.Contains("22522 CVC"))
                {
                    handicapCites.Add(citation);
                }
                else if (citation.location.Contains("N/OF") || citation.location.Contains("S/OF") || citation.location.Contains("E/OF") || citation.location.Contains("W/OF") || citation.location.Contains("A/F") || citation.location.Contains("MUNICIPAL LOT"))
                {
                    oldTownCites.Add(citation);
                }
                else
                {
                    otherCites.Add(citation);
                    DataTicketCitation.geoZone geoZone = citation.GetGeoZone();
                    if (geoZone == DataTicketCitation.geoZone.oldTown)
                    {
                        oldTownCites.Add(citation);
                    } 
                    else if (geoZone == DataTicketCitation.geoZone.northEnd)
                    {
                        northEndCites.Add(citation);
                    }
                    else if (geoZone == DataTicketCitation.geoZone.theHill)
                    {
                        theHillCites.Add(citation);
                    }
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
            button1.Visible = false;
            importDTData(fileName);
            button2.Visible = true;
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
                        otherCites.AddRange(handicapCites); // Append handiCapCites to otherCites
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
                }
                else if (firstWorksheet.Name == "Sample")
                {
                    StringBuilder test = new StringBuilder();
                    for (int row = start.Row + 6; row <= 52; row++)
                    {
                        if (!firstWorksheet.Cells[row, 2].Text.Contains("Count")) continue;
                        
                        int citeNumber = 0;
                        // OLDTOWN CITES
                        foreach (DataTicketCitation oldTownCite in oldTownCites)
                        {
                            string[] convDateTime = oldTownCite.dateTime.Split(" ");
                            string day = convDateTime[0];
                            firstWorksheet.Cells[row, 1].Style.Numberformat.Format = "mm/d/yy";
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 5].Value = citeNumber;
                        test.AppendLine(citeNumber.ToString());
                        citeNumber = 0;

                        // MAIN ST CITES
                        foreach (DataTicketCitation mainStCite in mainStCites)
                        {
                            string[] convDateTime = mainStCite.dateTime.Split(" ");
                            string day = convDateTime[0];
                            firstWorksheet.Cells[row, 1].Style.Numberformat.Format = "mm/d/yy";
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 6].Value = citeNumber;
                        citeNumber = 0;

                        // BEACH LOTS CITES
                        foreach (DataTicketCitation beachLotCite in beachLotCites)
                        {
                            string[] convDateTime = beachLotCite.dateTime.Split(" ");
                            string day = convDateTime[0];
                            firstWorksheet.Cells[row, 1].Style.Numberformat.Format = "mm/d/yy";
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 7].Value = citeNumber;
                        citeNumber = 0;


                        // NORTH END CITES
                        foreach (DataTicketCitation northEndCite in northEndCites)
                        {
                            string[] convDateTime = northEndCite.dateTime.Split(" ");
                            string day = convDateTime[0];
                            firstWorksheet.Cells[row, 1].Style.Numberformat.Format = "mm/d/yy";
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 8].Value = citeNumber;
                        citeNumber = 0;


                        // THE HILL CITES
                        foreach (DataTicketCitation theHillCite in theHillCites)
                        {
                            string[] convDateTime = theHillCite.dateTime.Split(" ");
                            string day = convDateTime[0];
                            firstWorksheet.Cells[row, 1].Style.Numberformat.Format = "mm/d/yy";
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 9].Value = citeNumber;
                        citeNumber = 0;


                        // HANDICAP CITES
                        foreach (DataTicketCitation handicapCite in handicapCites)
                        {
                            string[] convDateTime = handicapCite.dateTime.Split(" ");
                            string day = convDateTime[0];
                            firstWorksheet.Cells[row, 1].Style.Numberformat.Format = "mm/d/yy";
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 10].Value = citeNumber;
                        citeNumber = 0;


                        // STREET SWEEPER CITES
                        foreach (DataTicketCitation streetSweeperCite in sweeperCites)
                        {
                            string[] convDateTime = streetSweeperCite.dateTime.Split(" ");
                            string day = convDateTime[0];
                            firstWorksheet.Cells[row, 1].Style.Numberformat.Format = "mm/d/yy";
                            string worksheetDay = firstWorksheet.Cells[row, 1].Text;
                            if (day == worksheetDay)
                            {
                                citeNumber++;
                            }
                        }
                        firstWorksheet.Cells[row, 11].Value = citeNumber;
                        citeNumber = 0;
                    }
                    richTextBox1.Text = test.ToString();
                }
                excelPackage.Save();
            }
        }
    }
}
