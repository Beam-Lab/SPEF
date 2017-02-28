using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace BeamLab.SPEF
{
    public class SPEFContentTypeAttribute : Attribute
    {
        public SPEFContentTypeAttribute()
        {
        }
        public string Name { get; set; }
        public string Description { get; set; }
        
    }
}
