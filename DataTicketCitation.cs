using System;
using System.Collections.Generic;
using System.Text;

namespace SBPD_DT_Sentinel
{
    public class DataTicketCitation
    {
        public string dateTime { get; set; }
        public string citationNumber { get; set; }
        public string violationCode { get; set; }
        public string location { get; set; }
        public string comments { get; set; }

        public DataTicketCitation(string dateTime, string citationNumber, string violationCode, string location, string comments)
        {
            this.dateTime = dateTime;
            this.citationNumber = citationNumber;
            this.violationCode = violationCode;
            this.location = location;
            this.comments = comments;
        }
    }
}
