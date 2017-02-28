using Microsoft.SharePoint.Client;
using BeamLab.SPEF.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF.Models
{
    internal class SPEFStructInfo
    {
        public SPEFStructInfo()
        {
            Title = string.Empty;
            Description = string.Empty;
            FieldsInfo = new List<SPEFFieldInfo>();
            TemplateType = ListTemplateType.GenericList;
            ContentTypeName = string.Empty;
            BaseContentTypeName = string.Empty;
            SubSiteLabel = string.Empty;
            MainVariationLabel = string.Empty;
            UseVariations = false;
        }
        public bool IsList { get; set; }
        public bool IsContentType { get; set; }

        public string Title { get; set; }
        public string Description { get; set; }

        public string ContentTypeName { get; set; }
        public string ContentTypeID { get; set; }

        public string BaseContentTypeName { get; set; }

        public string SubSiteLabel { get; set; }
        public string SubSiteRelPath { get { return string.IsNullOrWhiteSpace(SubSiteLabel) ? string.Empty : string.Format("/{0}", SubSiteLabel); } }

        public bool UseVariations { get; set; }
        public string MainVariationLabel { get; set; }
        public string MainVariationRelPath { get { return string.IsNullOrWhiteSpace(MainVariationLabel) ? string.Empty : string.Format("/{0}", MainVariationLabel); } }

        public string ContextUrl { get { return string.Format("{0}{1}", SubSiteRelPath, MainVariationRelPath); } }

        public Type StructType { get; set; }

        public ListTemplateType TemplateType { get; set; }

        public List<SPEFFieldInfo> FieldsInfo { get; set; }
    }
}
