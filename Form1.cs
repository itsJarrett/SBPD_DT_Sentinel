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

        List<CADEvent> cadEvents = new List<CADEvent>();

        List<CADEvent> radioCalls = new List<CADEvent>();
        List<CADEvent> unitDetails = new List<CADEvent>();
        List<CADEvent> crimeReports = new List<CADEvent>();
        List<CADEvent> otherReports = new List<CADEvent>();
        List<CADEvent> indiaStorages = new List<CADEvent>();
        List<CADEvent> kiloStorages = new List<CADEvent>();
        List<CADEvent> oscarStorages = new List<CADEvent>();
        List<CADEvent> otherStorages  = new List<CADEvent>();
        List<CADEvent> radioBookings = new List<CADEvent>();
        List<CADEvent> radioTransports = new List<CADEvent>();

        public Form1()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            InitializeComponent();
        }

        private void ImportDTData(string fileName)
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

        private void ImportCADData(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];
                var start = firstWorksheet.Dimension.Start;
                var end = firstWorksheet.Dimension.End;
                for (int row = start.Row + 12; row <= end.Row - 5; row++)
                {
                    string eventNature = firstWorksheet.Cells[row, 1].Text;
                    string eventDate = firstWorksheet.Cells[row, 6].Text;
                    string eventLocation = firstWorksheet.Cells[row, 26].Text;
                    string eventReportNumber = firstWorksheet.Cells[row, 30].Text;
                    string eventEventNumber = firstWorksheet.Cells[row, 34].Text;
                    string eventDisposition = firstWorksheet.Cells[row, 38].Text;
                    CADEvent newCadEvent = new CADEvent(eventNature, eventDate, eventLocation, eventReportNumber, eventEventNumber, eventDisposition);
                    cadEvents.Add(newCadEvent);
                    Console.WriteLine(eventNature + " " + eventDate + " " + eventLocation + " " + eventReportNumber + " " + eventEventNumber + " " + eventDisposition);
                }
            }

            foreach (CADEvent cadEvent in cadEvents)
            {
                if (cadEvent.isCall)
                {
                    radioCalls.Add(cadEvent);
                }
                if (cadEvent.isDetail())
                {
                    unitDetails.Add(cadEvent);
                }
                if (cadEvent.isReport() && !cadEvent.isReportCategoryOther())
                {
                    crimeReports.Add(cadEvent);
                }
                if (cadEvent.isReport() && cadEvent.isReportCategoryOther())
                {
                    otherReports.Add(cadEvent);
                }
                if (cadEvent.isVehicleStorageReport() && cadEvent.GetVehicleStorageAuthoritySection() == CADEvent.vehicleStorageAuthoritySection.India)
                {
                    indiaStorages.Add(cadEvent);
                }
                if (cadEvent.isVehicleStorageReport() && cadEvent.GetVehicleStorageAuthoritySection() == CADEvent.vehicleStorageAuthoritySection.Oscar)
                {
                    oscarStorages.Add(cadEvent);
                }
                if (cadEvent.isVehicleStorageReport() && cadEvent.GetVehicleStorageAuthoritySection() == CADEvent.vehicleStorageAuthoritySection.Kilo)
                {
                    kiloStorages.Add(cadEvent);
                }
                if (cadEvent.isVehicleStorageReport() && cadEvent.GetVehicleStorageAuthoritySection() == CADEvent.vehicleStorageAuthoritySection.Other)
                {
                    otherStorages.Add(cadEvent);
                }
                if (cadEvent.isJailBooking()) 
                {
                    radioBookings.Add(cadEvent);
                } if (cadEvent.isTransport())
                {
                    radioTransports.Add(cadEvent);
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
            richTextBox1.Text = richTextBox1.Text + "\nPerforming citation analysis and Seal Beach mapping sector analysis... Program may take up to 5 minutes based on the amount of citations being parsed.\n\n STANDBY. BUTTON WILL APPEAR WHEN DONE!";
            ImportDTData(fileName);
            button2.Visible = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialogDT.FileName = "CAD Events Data"; // Default file name
            openFileDialogDT.DefaultExt = ".xls"; // Default file extension
            openFileDialogDT.Filter = "Excel files (.xlsx)|*.xlsx"; // Filter files by extension
            openFileDialogDT.ShowDialog();
            label3.Visible = true;
            string fileName = openFileDialogDT.FileName;
            button3.Visible = false;
            richTextBox1.Text = richTextBox1.Text + "\nImported CAD Event Data!";
            ImportCADData(fileName);
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

        public void processCADEventAmount(List<CADEvent> cadEventList, ExcelWorksheet excelWorksheet, int excelRow, int excelColumn)
        {
            int eventAmount = 0;
            foreach (CADEvent currentEvent in cadEventList)
            {
                string[] convDateTime = currentEvent.date.Split("/");
                string day = convDateTime[1];
                string worksheetDay = excelWorksheet.Cells[excelRow, 1].Text;
                if (day.Length == 1)
                    day = "0" + day;
                if (excelWorksheet.Name != "CSO MONTHLY")
                {
                    day = currentEvent.date;
                    excelWorksheet.Cells[excelRow, 1].Style.Numberformat.Format = "mm/dd/yyyy";
                    worksheetDay = excelWorksheet.Cells[excelRow, 1].Text;
                }
                Console.WriteLine(day + "   " + worksheetDay);
                if (day == worksheetDay)
                {
                    eventAmount++;
                }
            }
            excelWorksheet.Cells[excelRow, excelColumn].Value = eventAmount;
            if (excelWorksheet.Name != "CSO MONTHLY")
            {
                excelWorksheet.Cells[excelRow, 1].Style.Numberformat.Format = "m/d/yy";
            }
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
                    richTextBox1.Text = richTextBox1.Text + "\nMonthly Log recognized as field log... appending...";
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

                        // RADIO CALLS
                        processCADEventAmount(radioCalls, firstWorksheet, row, 2);

                        // DETAILS
                        processCADEventAmount(unitDetails, firstWorksheet, row, 3);

                        // CRIME REPORTS
                        processCADEventAmount(crimeReports, firstWorksheet, row, 4);

                        // OTHER REPORTS
                        processCADEventAmount(otherReports, firstWorksheet, row, 5);

                        // INDIA STORAGES
                        processCADEventAmount(indiaStorages, firstWorksheet, row, 7);

                        // KILO STORAGES
                        processCADEventAmount(kiloStorages, firstWorksheet, row, 8);

                        // OSCAR STORAGES
                        processCADEventAmount(oscarStorages, firstWorksheet, row, 9);

                        // OTHER STORAGES
                        processCADEventAmount(otherStorages, firstWorksheet, row, 10);
                    }
                }
                else if (firstWorksheet.Name == "Sample")
                {
                    richTextBox1.Text = richTextBox1.Text + "\nMonthly Log recognized as daily log... appending...";
                    for (int row = start.Row + 6; row <= 52; row++)
                    {
                        if (!firstWorksheet.Cells[row, 2].Text.Contains("Count")) continue;
                        
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

                        // BOOKINGS
                        processCADEventAmount(radioBookings, firstWorksheet, row, 13);

                        // TRANSPORTS
                        processCADEventAmount(radioTransports, firstWorksheet, row, 14);

                        // REPORTS
                        var totalReports = crimeReports.Concat(otherReports).Concat(indiaStorages).Concat(oscarStorages).Concat(kiloStorages).Concat(otherStorages).ToList();
                        processCADEventAmount(totalReports, firstWorksheet, row, 15);
                    }
                }
                excelPackage.Save();
            }
        }
    }
}
