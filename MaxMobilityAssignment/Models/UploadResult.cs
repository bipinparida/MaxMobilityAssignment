using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MaxMobilityAssignment.Models
{
    public class UploadResult
    {
        public int Row { get; set; }
        public string Status { get; set; }
        public int SerialNo { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string PhoneNo { get; set; }
        public string Address { get; set; }
    }
}