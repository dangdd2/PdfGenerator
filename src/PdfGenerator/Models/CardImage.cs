using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfGenerator
{
    public partial class CardImage
    {
        public CardImage()
        {
        }

        public int CardImageID { get; set; }
        public int CardID { get; set; }
        public string ImageData { get; set; }
        public DateTimeOffset CreatedDate { get; set; }
        public bool CardFront { get; set; }
    }
}
