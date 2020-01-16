using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace PdfGenerator
{
    public partial class Case
    {
        public Case()
        {
            this.Cards = new List<Card>();
        }

        public int CaseID { get; set; }
        public int ClientID { get; set; }
        public string CaseStatusID { get; set; }
        public string CaseNumber { get; set; }
        public int AssignedUserID { get; set; }
        public int CreatedUserID { get; set; }
        public DateTimeOffset CreatedDate { get; set; }
        public int ModifiedUserID { get; set; }
        public DateTimeOffset ModifiedDate { get; set; }
        public int? ClientGroupID { get; set; }
        public bool Locked { get; set; }
        [Display(Name = "Enable BOLO Notification and Search")]
        public bool INSEnabled { get; set; }
        public string NotificationTypeID { get; set; }

        public virtual List<Card> Cards { get; set; }

        public virtual Client Client { get; set; }

        public virtual User AssignedUser { get; set; }

        public virtual User CreatedUser { get; set; }

        public virtual User ModifiedUser { get; set; }
        public virtual ClientGroup ClientGroup { get; set; }
    }
}
