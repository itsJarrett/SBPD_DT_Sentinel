using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace SBPD_DT_Sentinel
{
    public class CADEvent
    {
        public string nature { get; set; }
        public string date { get; set; }
        public string location { get; set; }
        public string reportNumber { get; set; }
        public string eventnumber { get; set; }
        public string disposition { get; set; }
        public bool isCall { get; set; } = true;
        public enum vehicleStorageAuthoritySection
        {
            Oscar,
            India,
            Kilo,
            Other
        }

        public CADEvent(string nature, string date, string location, string reportNumber, string eventnumber, string disposition)
        {
            this.nature = nature;
            this.date = date;
            this.location = location;
            this.reportNumber = reportNumber;
            this.eventnumber = eventnumber;
            this.disposition = disposition;
            if (this.eventnumber == "")
                isCall = false;
        }

        public bool isVehicleStorageReport()
        {
            if (this.disposition.Contains("SVS") || (this.disposition.Contains("RPT") && (this.nature.Contains("926") || this.nature.Contains("586") || this.nature.Contains("917A"))))
                return true;
            return false;
        }

        public vehicleStorageAuthoritySection GetVehicleStorageAuthoritySection() {
            if (this.nature.Contains("917A")) {
                return vehicleStorageAuthoritySection.Kilo;
            }
            if (this.nature.Contains("VEHCK")) {
                return vehicleStorageAuthoritySection.Oscar;
            }
            if (this.nature.Contains("926"))
            {
                return vehicleStorageAuthoritySection.India;
            }
            return vehicleStorageAuthoritySection.Other;
        }

        public bool isReport()
        {
            if (this.disposition.Contains("SVS") || (this.disposition.Contains("RPT")))
                return true;
            return false;
        }

        public bool isReportCategoryOther()
        {
            if (this.nature.Contains("928") || this.nature == "LP")
                return true;
            return false;
        }

        public bool isJailBooking()
        {
            if (this.nature.Contains("JB"))
                return true;
            return false;
        }

        public bool isTransport()
        {
            if (this.nature.Contains("TP"))
                return true;
            return false;
        }

        public bool isDetail()
        {
            if (this.nature == "924")
                return true;
            return false;
        }
    }
}
