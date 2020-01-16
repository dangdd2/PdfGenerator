using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace PdfGenerator
{
    public partial class ClientGroup
    {
        public ClientGroup()
        {
            this.SubGroups = new List<ClientGroup>();
            this.ClientGroupUsers = new List<User>();
            this.ClientGroupCases = new List<Case>();
        }

        public int ClientGroupID { get; set; }
        public int ClientID { get; set; }

        [Display(Name = "Group")]
        public string Name { get; set; }
        public int? ParentClientGroupID { get; set; }
        public int? ClientMerchantAccountId { get; set; }
        public virtual Client Client { get; set; }
        public virtual ICollection<ClientGroup> SubGroups { get; set; }
        public virtual ICollection<User> ClientGroupUsers { get; set; }
        public virtual ICollection<Case> ClientGroupCases { get; set; }

        public virtual ClientGroup ParentGroup { get; set; }
    }
}
