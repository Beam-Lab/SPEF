using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF
{
    public class SPEFUser
    {
        public SPEFUser()
        {
            IDs = new List<int>();
            Groups = new List<SPEFUser>();
            IsGroup = false;
            Members = new List<SPEFUser>();
            Properties = new Dictionary<string, string>();
        }

        public SPEFUser(int id)
        {
            IDs = new List<int>();
            ID = id;
            Groups = new List<SPEFUser>();
            IsGroup = false;
            Members = new List<SPEFUser>();
            Properties = new Dictionary<string, string>();
        }

        private int id;
        public int ID {
            get
            {
                return id;
            }
            set
            {
                if (!IDs.Contains(value))
                    IDs.Add(value);
                id = value;
            } 
        }
        public List<int> IDs { get; set; }
        public string AccountName { get; set; }
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public List<SPEFUser> Groups { get; set; }
        public string UserUrl { get; set; }
        public string PictureUrl { get; set; }
        public bool IsSiteAdmin { get; set; }

        public IDictionary<string, string> Properties { get; set; }

        public bool IsGroup { get; set; }
        public List<SPEFUser> Members { get; set; }
    }

    //public class SPEFGroup : SPEFUser
    //{
    //    public SPEFGroup()
    //    {
    //        Members = new List<SPEFUser>();
    //    }

    //    public List<SPEFUser> Members { get; set; }
    //}
}
