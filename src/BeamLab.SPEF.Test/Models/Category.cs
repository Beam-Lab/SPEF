using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF.Test.Models
{
    [SPEFList(Title = "News Category", Description = "Categorie delle News")]
    public class Category : SPEFListItem
    {
        [SPEFNumericField]
        public int Position { get; set; }

        [SPEFUrlField]
        public string ImageUrl { get; set; }
    }
}
