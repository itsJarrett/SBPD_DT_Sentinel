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
                else if (citation.violationCode.Contains("22507.8") || citation.violationCode.Contains("22522"))
                {
                    handicapCites.Add(citation);
                    otherCites.Add(citation);
                }
                else if (citation.location.Contains("N/OF") || citation.location.Contains("S/OF") || citation.location.Contains("E/OF") || citation.location.Contains("W/OF") || citation.location.Contains("A/F") || citation.location.Contains("MUNICIPAL LOT"))
                {
                    oldTownCites.Add(citation);
                    otherCites.Add(citation);
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
            richTextBox1.Text = "Performing citation analysis and Seal Beach mapping sector analysis... Program may take up to 5 minutes based on the amount of citations being parsed.\n\n STANDBY. BUTTON WILL APPEAR WHEN DONE!";
            importDTData(fileName);
            button2.Visible = true;
        }

        public void processCitationAmount(List<DataTicketCitation> citationList, ExcelWorksheet excelWorksheet, int excelRow, int excelColumn)
        {
            int citationAmount = 0;
            foreach (DataTicketCitation currentCitation in citationList)
            {
                string[] convDateTime = currentCitation.dateTime.Split("/");
                string day = convDateTime[1];
                string worksheetDay = excelWorksheet.Cells[excelRow, 1].Text;
                if (day.Length == 1)
                    day = "0" + day;
                if (excelWorksheet.Name != "CSO MONTHLY")
                {
                    convDateTime = currentCitation.dateTime.Split(" ");
                    day = convDateTime[0];
                    excelWorksheet.Cells[excelRow, 1].Style.Numberformat.Format = "m/d/yy";
                    worksheetDay = excelWorksheet.Cells[excelRow, 1].Text;
                }
                if (day == worksheetDay)
                {
                    citationAmount++;
                }
            }
            excelWorksheet.Cells[excelRow, excelColumn].Value = citationAmount;
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
                    richTextBox1.Text = richTextBox1.Text + "\n\nMonthly Log recognized as field log... appending...";
                    for (int row = start.Row + 6; row <= 37; row++)
                    {
                        // ONE HOUR CITES
                        processCitationAmount(oneHourCites, firstWorksheet, row, 12);

                        // SWEEPER CITES
                        processCitationAmount(sweeperCites, firstWorksheet, row, 13);

                        // MAIN ST CITES
                        processCitationAmount(mainStCites, firstWorksheet, row, 14);

                        // BEACH LOT CITES
                        processCitationAmount(beachLotCites, firstWorksheet, row, 15);

                        // OTHER CITES
                        processCitationAmount(otherCites, firstWorksheet, row, 16);
                    }
                }
                else if (firstWorksheet.Name == "Sample")
                {
                    richTextBox1.Text = richTextBox1.Text + "\n\nMonthly Log recognized as daily log... appending...";
                    for (int row = start.Row + 6; row <= 52; row++)
                    {
                        if (!firstWorksheet.Cells[row, 2].Text.Contains("Count")) continue;
                        
                        int citeNumber = 0;
                        // OLDTOWN CITES
                        processCitationAmount(oldTownCites, firstWorksheet, row, 5);

                        // MAIN ST CITES
                        processCitationAmount(mainStCites, firstWorksheet, row, 6);

                        // BEACH LOTS CITES
                        processCitationAmount(beachLotCites, firstWorksheet, row, 7);

                        // NORTH END CITES
                        processCitationAmount(northEndCites, firstWorksheet, row, 8);

                        // THE HILL CITES
                        processCitationAmount(theHillCites, firstWorksheet, row, 9);


                        // HANDICAP CITES
                        processCitationAmount(handicapCites, firstWorksheet, row, 10);

                        // STREET SWEEPER CITES
                        processCitationAmount(sweeperCites, firstWorksheet, row, 11);
                    }
                }
                excelPackage.Save();
            }
        }
    }
}
