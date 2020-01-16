using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PdfGenerator
{
    public class CardMetaDataVM
    {
        public string Brand { get; set; }
        public string TypeOfCard { get; set; }
        public string Category { get; set; }
        public string Country { get; set; }
        public string Website { get; set; }
        public string IssuerName { get; set; }
        public string PhoneNumber { get; set; }
        public string Address { get; set; }
    }
}