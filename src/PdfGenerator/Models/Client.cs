using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PdfGenerator
{
    public partial class Client
    {
        [Flags]
        public enum INSPreferenceOptions : int
        {
            None = 0,
            Credit = 1,
            Debit = 2,
            Prepaid = 4
        }

        public Client()
        {
            this.Cases = new List<Case>();
            this.Users = new List<User>();
            this.ClientGroups = new List<ClientGroup>();
        }

        public int ClientID { get; set; }
        [Display(Name = "Client")]
        public string Name { get; set; }
        public string Login { get; set; }
        public bool Active { get; set; }
        public string Phone { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zipcode { get; set; }
        public string TimeZoneID { get; set; }
        public string PaymentProcessor { get; set; }
        public string PPAPIKey { get; set; }
        public string PPClientID { get; set; }
        public string PPUsername { get; set; }
        public string PPPassword { get; set; }
        public int CreatedUserID { get; set; }
        public int? ParentClientID { get; set; }
        public bool AllowCardSearch { get; set; }
        
        public string ClientTypeID { get; set; }

        public string EPXCustomerNbr { get; set; }
        public string EPXMerchantNbr { get; set; }
        public string EPXDbaNbr { get; set; }
        public string EPXTerminalNbr { get; set; }

        public int? RequiredAgreementID { get; set; }

        public int? DefaultMerchantAccountId { get; set; }

        public INSPreferenceOptions INSPreferences { get; set; }

        public string ReferralSource { get; set; }
        public string SalesRep { get; set; }
        public DateTime? ContractEffectiveDate { get; set; }
        public DateTime? ContractBillingRenewalDate { get; set; }
        public DateTime? ContractExpirationDate { get; set; }
        public string BillingType { get; set; }
        public double BillingRate{ get; set; }
        public double UserSubscriptionRate { get; set; }
        public int? AuthUsersLimit { get; set; }

        public string AdminUsername { get; set; }
        public string AdminPassword { get; set; }
        public string Product { get; set; }
        public string BillingMethod { get; set; }

        [Display(Name = "Custom 1")]
        [MaxLength(200)]
        public string UDef1 { get; set; }
        [Display(Name = "Custom 2")]
        [MaxLength(200)]
        public string UDef2 { get; set; }
        [Display(Name = "Custom 3")]
        [MaxLength(200)]
        public string UDef3 { get; set; }


        public DateTimeOffset CreatedDate { get; set; }
        public virtual ICollection<Case> Cases { get; set; }
        public virtual ICollection<User> Users { get; set; }
        public virtual User CreatedUser { get; set; }
        public virtual ICollection<ClientGroup> ClientGroups { get; set; }
        public virtual Client ParentClient { get; set; }
        public virtual ICollection<Client> NestedClients { get; set; }
    }
}
