using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PdfGenerator
{
    public partial class User
    {
        public User()
        {
            this.Cards = new List<Card>();
            this.AssignedCases = new List<Case>();
            this.CreatedCases = new List<Case>();
            this.ModifiedCases = new List<Case>();
            this.CreatedClients = new List<Client>();
            this.CreatedDate = DateTime.Now;
        }
        public int UserID { get; set; }
        [EmailAddress]
        public string Email { get; set; }
        public string PasswordHash { get; set; }
        [Display(Name="Last Name")]
        public string LastName { get; set; }
        [Display(Name = "First Name")]
        public string FirstName { get; set; }
        public bool Active { get; set; }
        public string AuthType { get; set; }
        public string UserRoleID { get; set; }
        public int ClientID { get; set; }
        public int? ClientGroupID { get; set; }
        public int? AdministrativeGroupID { get; set; }
        public DateTimeOffset CreatedDate { get; set; }

        public int? LastAgreementID { get; set; }
        public Nullable<DateTimeOffset> LastLoginDate { get; set; }
        public Nullable<DateTimeOffset> LastAcceptDate { get; set; }
        public Nullable<DateTimeOffset> LastResetDate { get; set; }
        public virtual ICollection<Card> Cards { get; set; }
        public virtual ICollection<Case> AssignedCases { get; set; }
        public virtual ICollection<Case> CreatedCases { get; set; }
        public virtual ICollection<Case> ModifiedCases { get; set; }
        public virtual ICollection<Client> CreatedClients { get; set; }
        public virtual Client Client { get; set; }
        public virtual ClientGroup ClientGroup { get; set; }

        public string FriendlyName
        {
            get
            {
                if (FirstName != null && FirstName.Length > 0)
                {
                    return (FirstName + " " + LastName).Trim();
                }

                return "";
            }
        }
        public string DisplayName
        {
            get
            {
                if (LastName != null && LastName.Length > 0)
                {
                    return (LastName + ", " + FirstName).Trim();
                }

                return "";
            }
        }
    }
}
