using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF
{
    public class SPEFTaxonomyItem
    {
        public SPEFTaxonomyItem()
        {
            CustomProperties = new Dictionary<string, string>();
        }
        public Guid ID { get; set; }
        public string Value { get; set; }

        public Dictionary<string, string> CustomProperties { get; set; }
    }
}
