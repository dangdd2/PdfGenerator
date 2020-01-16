using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.RegularExpressions;

namespace PdfGenerator
{
    public partial class Card
    {
        public Card()
        {
            this.CardTransactions = new List<CardTransaction>();
			this.ClonedCards = new List<Card>();
            this.CardImages = new List<CardImage>();
        }
        public int CardID { get; set; }
        public int CaseID { get; set; }
        public int? ProviderID { get; set; }
        public string CardNumberLast4 { get; set; }
        public string Notes { get; set; }

        public Nullable<decimal> CardBalance { get; set; }
        public string BalanceSource { get; set; } //Manual, System
        public Nullable<decimal> OriginalBalance { get; set; }
        public string OriginalBalanceSource { get; set; } //Manual, System
        public Nullable<System.DateTimeOffset> CardBalanceDate { get; set; }
        public string Status { get; set; }
        public string CardNumberUnencrypted { get; set; }
        public string CardNumberHash { get; set; }
        public string CardNumberMasked { get; set; }
        public string CardType { get; set; }
        public string BinBaseBIN { get; set; }
        public string CardName { get; set; }
        public string CardExpiration { get; set; }
        public string IssuerName { get; set; }
        public string IssuerSource { get; set; }
        public string RetailerName { get; set; } //For cards with "Other" provider
        public byte[] CardCVV { get; set; }
        public string EntryMethod { get; set; } //Swipe or Manual
        public string TrackData { get; set; }
        public int CreatedUserID { get; set; }
        public DateTimeOffset CreatedDate { get; set; }
        public bool MissingPrintedInfo { get; set; }
		public int? ClonedCardParentCardId { get; set; }
        //User-Defined Columns
        [MaxLength(2000)]
        public string UDef1 { get; set; }
        [MaxLength(2000)]
        public string UDef2 { get; set; }
        [MaxLength(2000)]
        public string UDef3 { get; set; }

        public virtual Case Case { get; set; }
        public virtual Provider Provider { get; set; }
        public virtual User User { get; set; }
		public virtual BinBase BinBase { get; set; }
        public virtual ICollection<CardTransaction> CardTransactions { get; set; }
		public virtual Card ClonedCardParent{ get; set; }
		public virtual List<Card> ClonedCards { get; set; }
        public virtual List<CardImage> CardImages { get; set; }

	}
}
