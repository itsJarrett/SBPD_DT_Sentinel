using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
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

        public enum geoZone
        {
            oldTown,
            northEnd,
            theHill
        }

        public DataTicketCitation(string dateTime, string citationNumber, string violationCode, string location, string comments)
        {
            this.dateTime = dateTime;
            this.citationNumber = citationNumber;
            this.violationCode = violationCode;
            this.location = location;
            this.comments = comments;
        }

        public geoZone GetGeoZone()
        {
            Dictionary<geoZone, double> geoZoneDistances = new Dictionary<geoZone, double>();
            using (WebClient wc = new WebClient())
            {
                string bingLocation = location + ", Seal Beach CA 90740";
                var json = wc.DownloadString("http://dev.virtualearth.net/REST/V1/Routes/Walking?wp.0=" + bingLocation + "&wp.1=" + "215 Main St, Seal Beach CA 90740" + "&key=KqIPiiPICvJvqRalm1Y9~yDGiUeoW_JI27wt0hI3x6w~Aqe8NJcoZIoyYT_dVxh61HQLUN10dVMO-YQXEF6o_Nj21Vhqq0GGnmTZuBd6BatK");
                dynamic jsonObj = JsonConvert.DeserializeObject(json);
                double travelDistance = jsonObj.resourceSets[0].resources[0].travelDistance;
                geoZoneDistances.Add(geoZone.oldTown, travelDistance);

                json = wc.DownloadString("http://dev.virtualearth.net/REST/V1/Routes/Walking?wp.0=" + bingLocation + "&wp.1=" + "13900 Seal Beach Blvd, Seal Beach CA 90740" + "&key=KqIPiiPICvJvqRalm1Y9~yDGiUeoW_JI27wt0hI3x6w~Aqe8NJcoZIoyYT_dVxh61HQLUN10dVMO-YQXEF6o_Nj21Vhqq0GGnmTZuBd6BatK");
                jsonObj = JsonConvert.DeserializeObject(json);
                travelDistance = jsonObj.resourceSets[0].resources[0].travelDistance;
                geoZoneDistances.Add(geoZone.northEnd, travelDistance);

                json = wc.DownloadString("http://dev.virtualearth.net/REST/V1/Routes/Walking?wp.0=" + bingLocation + "&wp.1=" + "1430 Catalina Ave, Seal Beach CA 90740" + "&key=KqIPiiPICvJvqRalm1Y9~yDGiUeoW_JI27wt0hI3x6w~Aqe8NJcoZIoyYT_dVxh61HQLUN10dVMO-YQXEF6o_Nj21Vhqq0GGnmTZuBd6BatK");
                jsonObj = JsonConvert.DeserializeObject(json);
                travelDistance = jsonObj.resourceSets[0].resources[0].travelDistance;
                geoZoneDistances.Add(geoZone.theHill, travelDistance);
            }
            return geoZoneDistances.Aggregate((l, r) => l.Value < r.Value ? l : r).Key;
        }
    }
}
