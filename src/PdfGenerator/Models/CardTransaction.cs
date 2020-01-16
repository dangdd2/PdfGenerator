using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfGenerator
{
    public partial class CardTransaction
    {
        public CardTransaction()
        {
        }

        public int CardTransactionID { get; set; }
        public int CardID { get; set; }
        public string TransactionType { get; set; }
        public string Status { get; set; }
        public string Message { get; set; }
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:0.00}")]
        public Nullable<decimal> Amount { get; set; }
        public string ExternalTransactionID { get; set; }
        public string TransactionDetails { get; set; }
        public Nullable<int> CaseBatchID { get; set; }
        public DateTimeOffset RequestedDate { get; set; }
        public int RequestedUserID { get; set; }
        public Nullable<DateTimeOffset> CompletedDate { get; set; }
        public string PaymentProcessor { get; set; }

        public int? ClientMerchantAccountId { get; set; }

        public virtual Card Card { get; set; }
        public virtual User RequestedUser { get; set; }
    }
}
