using System;
using System.Collections.Generic;

namespace PdfGenerator
{
    public partial class IAFCI
    {
        public int IAFCIID { get; set; }
        public string BIN { get; set; }
        public string Brand { get; set; }
        public string Type { get; set; }
        public string Category { get; set; }
        public string IssuerName { get; set; }
        public string Website { get; set; }
        public string PhoneNumber { get; set; }
        public string Address { get; set; }
        public string PhysicalAddress { get; set; }
        public DateTimeOffset CreatedDate { get; set; }
        public DateTimeOffset ModifiedDate { get; set; }
    }
}
