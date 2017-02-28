using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF.Models
{
    internal class SPEFFieldInfo
    {
        public PropertyInfo PropertyInfo
        {
            get;
            private set;
        }
        public SPEFFieldInfo(PropertyInfo propertyInfo)
        {
            PropertyInfo = propertyInfo;
            Ignore = false;
            Readonly = false;
            ShowInDisplayForm = true;
            ShowInEditForm = true;
            ShowInNewForm = true;
            ShowInViewForm = true;
            Choices = new List<string>();
            Type = FieldType.Text;
            AdditionalAttributes = new Dictionary<string, string>();
        }
        public bool Ignore { get; set; }
        public bool Readonly { get; set; }
        public Guid ID { get; set; }
        /// <summary>
        /// InternalName
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Title & DisplayName
        /// </summary>
        public string Title { get; set; }
        //[JsonConverter(typeof(StringEnumConverter))]
        public FieldType Type { get; set; }
        public int DecimalPlaces { get; set; }
        public bool FieldTypeTaxonomy { get; internal set; }
        public int MaxLength { get; set; }
        public bool Required { get; set; }
        public bool Multiple { get; set; }
        public bool ShowInDisplayForm { get; set; }
        public bool ShowInEditForm { get; set; }
        public bool ShowInNewForm { get; set; }
        public bool ShowInViewForm { get; set; }
        public bool Hidden { get; set; }
        public string DefaultValue { get; set; }
        public List<string> Choices { get; set; }
        public bool RichText { get; set; }
        //[JsonConverter(typeof(StringEnumConverter))]
        //public RichTextMode RichTextMode { get; set; }
        public string RichTextMode { get; set; }
        public DateTimeFieldFormatType DateTimeFormatType { get; set; }
        //[JsonConverter(typeof(StringEnumConverter))]
        public UrlFieldFormatType UrlFormatType { get; set; }
        //[JsonConverter(typeof(StringEnumConverter))]
        public FieldUserSelectionMode UserSelectionMode { get; set; }
        public string LookupList { get; set; }
        internal Guid LookupListID { get; set; }
        public string LookupField { get; set; }

        /// <summary>
        /// Taxonomy Field
        /// </summary>
        internal Guid TermStoreId { get; set; }
        internal Guid TermGroupId { get; set; }
        internal Guid TermSetId { get; set; }
        internal Guid TaxFieldId { get; set; }
        public string TermStoreName { get; set; }
        public string TermGroupName { get; set; }
        public string TermSetName { get; set; }
        public bool TermUserCreated { get; set; }
        public string TaxFieldName { get; set; }


        public bool Overwrite { get; set; }
        public Dictionary<string, string> AdditionalAttributes { get; set; }

        //[JsonIgnore]
        private string XmlFieldType
        {
            get
            {
                if (Type == FieldType.Lookup && FieldTypeTaxonomy)
                    return "TaxonomyFieldType";

                if (Multiple && (Type == FieldType.User || Type == FieldType.Lookup))
                    return string.Format("{0}Multi", Type.ToString());
                return Type.ToString();
            }
        }
        //[JsonIgnore]
        internal string XmlName
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(Name))
                    return WebUtility.HtmlEncode(Name);
                return WebUtility.HtmlEncode(Title);
            }
        }
        //[JsonIgnore]
        private string XmlDisplayName
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(Title))
                    return WebUtility.HtmlEncode(Title);
                return WebUtility.HtmlEncode(Name);
            }
        }
        public string GetXml()
        {
            var retXml = new StringBuilder();
            var baseXml = string.Format("<Field StaticName='{0}' Name='{0}' DisplayName='{1}' Type='{2}' Required='{3}' ShowInDisplayForm='{4}' ShowInEditForm='{5}' ShowInNewForm='{6}' ShowInViewForms='{7}' Hidden='{8}'",
                XmlName,
                XmlDisplayName,
                XmlFieldType,
                (Required ? "TRUE" : "FALSE"),
                (ShowInDisplayForm ? "TRUE" : "FALSE"),
                (ShowInEditForm ? "TRUE" : "FALSE"),
                (ShowInNewForm ? "TRUE" : "FALSE"),
                (ShowInViewForm ? "TRUE" : "FALSE"),
                (Hidden ? "TRUE" : "FALSE"));

            retXml.Append(baseXml);

            if (Overwrite)
            {
                retXml.Append(" Overwrite='TRUE'");
            }

            if (DecimalPlaces > 0 && (Type == FieldType.Currency || Type == FieldType.Number))
            {
                var decimalPlacesXml = string.Format(" Decimals='{0}'", DecimalPlaces);
                retXml.Append(decimalPlacesXml);
            }

            if (Type == FieldType.Text)
            {
                var maxLength = 255;
                if (MaxLength > 0 && MaxLength <= 255)
                    maxLength = MaxLength;
                var textXml = string.Format(" MaxLength='{0}'", maxLength);
                retXml.Append(textXml);
            }

            if (Type == FieldType.Note)
            {
                var noteXml = string.Format(" RichText='{0}'", (RichText ? "TRUE" : "FALSE"));
                retXml.Append(noteXml);

                if (RichText)
                {
                    var richTextXml = string.Format(" RichTextMode='{0}'", RichTextMode.ToString());
                    retXml.Append(richTextXml);
                }
            }

            if (Type == FieldType.DateTime)
            {
                var dateTimeXml = string.Format(" Format='{0}'", DateTimeFormatType.ToString());
                retXml.Append(dateTimeXml);
            }

            if (Type == FieldType.URL)
            {
                var urlXml = string.Format(" Format='{0}'", UrlFormatType.ToString());
                retXml.Append(urlXml);
            }

            if (Type == FieldType.Lookup)
            {
                if (!FieldTypeTaxonomy)
                {
                    var lookupXml = string.Format(" List='{0}' ShowField='{1}'", "{" + LookupListID + "}", LookupField);
                    retXml.Append(lookupXml);
                }
                else
                {
                    string taxField = string.Format(" ShowField='Term1033' EnforceUniqueValues='FALSE' Sortable='FALSE' Group='My Group'");
                }
            }

            if (Type == FieldType.User || Type == FieldType.Lookup)
            {
                if (Type == FieldType.User)
                {
                    var userXml = string.Format(" UserSelectionMode='{0}'", UserSelectionMode.ToString());
                    retXml.Append(userXml);
                }

                if (Multiple)
                    retXml.Append(" Mult='TRUE'");
            }

            foreach (var attribute in AdditionalAttributes)
            {
                retXml.AppendFormat(" {0}='{1}'", attribute.Key, attribute.Value);
            }

            retXml.Append(">");

            if (Type == FieldType.Lookup && FieldTypeTaxonomy)
            {
                string taxField0 = string.Format(@"<Customization><ArrayOfProperty>
                    <Property><Name>SspId</Name><Value xmlns:q1='http://www.w3.org/2001/XMLSchema' p4:type='q1:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{0}</Value></Property>
                    <Property><Name>GroupId</Name></Property><Property><Name>TermSetId</Name><Value xmlns:q2='http://www.w3.org/2001/XMLSchema' p4:type='q2:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{1}</Value></Property>
                    <Property><Name>AnchorId</Name><Value xmlns:q3='http://www.w3.org/2001/XMLSchema' p4:type='q3:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>00000000-0000-0000-0000-000000000000</Value></Property>
                    <Property><Name>UserCreated</Name><Value xmlns:q4='http://www.w3.org/2001/XMLSchema' p4:type='q4:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{2}</Value></Property>
                    <Property><Name>Open</Name><Value xmlns:q5='http://www.w3.org/2001/XMLSchema' p4:type='q5:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value></Property>
                    <Property><Name>TextField</Name><Value xmlns:q6='http://www.w3.org/2001/XMLSchema' p4:type='q6:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{3}</Value></Property>
                    <Property><Name>IsPathRendered</Name><Value xmlns:q7='http://www.w3.org/2001/XMLSchema' p4:type='q7:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>true</Value></Property>
                    <Property><Name>IsKeyword</Name><Value xmlns:q8='http://www.w3.org/2001/XMLSchema p4:type=q8:boolean xmlns:p4=http://www.w3.org/2001/XMLSchema-instance'>false</Value></Property>
                    <Property><Name>TargetTemplate</Name></Property><Property><Name>CreateValuesInEditForm</Name><Value xmlns:q9='http://www.w3.org/2001/XMLSchema' p4:type='q9:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value></Property>
                    <Property><Name>FilterAssemblyStrongName</Name><Value xmlns:q10='http://www.w3.org/2001/XMLSchema' p4:type='q10:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>Microsoft.SharePoint.Taxonomy, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Value></Property>
                    <Property><Name>FilterClassName</Name><Value xmlns:q11='http://www.w3.org/2001/XMLSchema' p4:type='q11:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>Microsoft.SharePoint.Taxonomy.TaxonomyField</Value></Property>
                    <Property><Name>FilterMethodName</Name><Value xmlns:q12='http://www.w3.org/2001/XMLSchema' p4:type='q12:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>GetFilteringHtml</Value></Property>
                    <Property><Name>FilterJavascriptProperty</Name><Value xmlns:q13='http://www.w3.org/2001/XMLSchema' p4:type='q13:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>FilteringJavascript</Value></Property>
                    </ArrayOfProperty></Customization>",
                    TermStoreId.ToString("D"),
                    TermSetId.ToString("D"),
                    TermUserCreated.ToString().ToLower(),
                    TaxFieldId.ToString("B")
                );
                string taxField = string.Format(@"<Customization><ArrayOfProperty>
                    <Property><Name>SspId</Name><Value xmlns:q1='http://www.w3.org/2001/XMLSchema' p4:type='q1:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{0}</Value></Property>
                    <Property><Name>TermSetId</Name><Value xmlns:q2='http://www.w3.org/2001/XMLSchema' p4:type='q2:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{1}</Value></Property>
                    <Property><Name>UserCreated</Name><Value xmlns:q4='http://www.w3.org/2001/XMLSchema' p4:type='q4:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{2}</Value></Property>
                    <Property><Name>TextField</Name><Value xmlns:q6='http://www.w3.org/2001/XMLSchema' p4:type='q6:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{3}</Value></Property>
                    <Property><Name>IsPathRendered</Name><Value xmlns:q7='http://www.w3.org/2001/XMLSchema' p4:type='q7:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>true</Value></Property>
                    <Property><Name>CreateValuesInEditForm</Name><Value xmlns:q9='http://www.w3.org/2001/XMLSchema' p4:type='q9:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value></Property>
                    </ArrayOfProperty></Customization>",
                   TermStoreId.ToString("D"),
                   TermSetId.ToString("D"),
                   TermUserCreated.ToString().ToLower(),
                   TaxFieldId.ToString("B")
               );
                retXml.Append(taxField);
            }

            //Proprietà nell'InnerXml
            if (Type == FieldType.Choice || Type == FieldType.MultiChoice)
            {
                var choices = Choices.Select(c => string.Format("<CHOICE>{0}</CHOICE>", c));
                var choicesXml = string.Format("<CHOICES>{0}</CHOICES>", string.Join("", choices.Select(c => c)));
                retXml.Append(choicesXml);
            }
            retXml.Append("</Field>");
            return retXml.ToString();
        }
    }
}
