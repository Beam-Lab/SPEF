using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace BeamLab.SPEF
{
    public class SPEFListAttribute : Attribute
    {
        public SPEFListAttribute()
        {
            TemplateType = ListTemplateType.GenericList;
            ContentTypeName = string.Empty;
            SubSiteLabel = string.Empty;
            UseVariations = false;
        }
        public string Title { get; set; }
        public string Description { get; set; }
        public ListTemplateType TemplateType { get; set; }

        public string ContentTypeName { get; set; }

        public string  SubSiteLabel { get; set; }
        public bool UseVariations { get; set; }
    }
}
