using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace PdfGenerator
{
    public partial class Provider
    {
        public Provider()
        {
            this.Cards = new List<Card>();
        }

        public int ProviderID { get; set; }
        public string Name { get; set; }
        public bool Enabled { get; set; }
        public string URL { get; set; }
        public bool CardNumberRequired { get; set; }
        public bool PinRequired { get; set; }
        public bool LoginRequired { get; set; }
        public string LoginURL { get; set; }
        public string LoginUsernameFieldType { get; set; }
        public string LoginUsernameFieldIdentifier { get; set; }
        public string LoginPasswordFieldType { get; set; }
        public string LoginPasswordFieldIdentifier { get; set; }
        public string LoginSubmitFieldType { get; set; }
        public string LoginSubmitFieldIdentifier { get; set; }
        public string CaptchaErrorFieldType { get; set; }
        public string CaptchaErrorFieldIdentifier { get; set; }
        public string CardNumberFieldType { get; set; }
        public string CardNumberFieldIdentifier { get; set; }
        public string CardPinFieldType { get; set; }
        public string CardPinFieldIdentifier { get; set; }
        public bool CombineCardPin { get; set; }
        public bool CaptchaRequired { get; set; }
        public bool RecaptchaRequired { get; set; }
        public string RecaptchaSiteKey { get; set; }
        public string CaptchaFieldType { get; set; }
        public string CaptchaFieldIdentifier { get; set; }
        public string CaptchaImageType { get; set; }
        public string CaptchaImageIdentifier { get; set; }
        public string SubmitFieldType { get; set; }
        public string SubmitFieldIdentifier { get; set; }
        public int SubmitWaitTime { get; set; }
        public string ErrorFieldType { get; set; }
        public string ErrorFieldIdentifier { get; set; }
        public string BalanceFieldType { get; set; }
        public string BalanceFieldIdentifier { get; set; }
        public string BalanceParseRegex { get; set; }
        public string CardImage { get; set; }
        public string Logo { get; set; }
        public string PreClickFieldType { get; set; }
        public string PreClickFieldIdentifier { get; set; }
        public string SwipePattern { get; set; }
        public bool Offline { get; set; }
        public int ModifiedUserID { get; set; }
        public DateTimeOffset ModifiedDate { get; set; }

        public string SubpoenaName { get; set; }
        public string SubpoenaAttention { get; set; }
        public string SubpoenaAddress1 { get; set; }
        public string SubpoenaAddress2 { get; set; }
        public string SubpoenaCity { get; set; }
        public string SubpoenaState { get; set; }
        public string SubpoenaZip { get; set; }
        public string SubpoenaFax { get; set; }
        public string SubpoenaPhone { get; set; }

        [MaxLength(50)]
        public string CorporatePhone { get; set; }

        public string CardBrand { get; set; }
        public string TypeOfCard { get; set; }
        public string Category { get; set; }

        public virtual ICollection<Card> Cards { get; set; }
        public virtual User ModifiedUser { get; set; }
    }
}
