using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF
{
    public abstract class SPEFFieldAttribute : Attribute
    {
        public SPEFFieldAttribute()
        {
            Ignore = false;
            Readonly = false;
            Name = string.Empty;
            Title = string.Empty;
            Hidden = false;
            Required = false;
            ShowInNew = true;
            ShowInDisplay = true;
            ShowInEdit = true;
            ShowInView = true;
        }
        /// <summary>
        /// Ignore on create/update Repository
        /// </summary>
        public bool Ignore { get; set; }
        /// <summary>
        /// Ignore on set value
        /// </summary>
        public bool Readonly { get; set; }
        /// <summary>
        /// Internal Name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Title & Display Name
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// Default: Text
        /// </summary>
        public abstract FieldType FieldType { get; }
        /// <summary>
        /// Default: true
        /// </summary>
        public bool Required { get; set; }
        /// <summary>
        /// Default: false
        /// </summary>
        public bool Hidden { get; set; }
        /// <summary>
        /// Default: true
        /// </summary>
        public bool ShowInNew { get; set; }
        /// <summary>
        /// Default: true
        /// </summary>
        public bool ShowInDisplay { get; set; }
        /// <summary>
        /// Default: true
        /// </summary>
        public bool ShowInEdit { get; set; }
        /// <summary>
        /// Default: true
        /// </summary>
        public bool ShowInView { get; set; }
    }

    public class SPEFNumericFieldAttribute : SPEFFieldAttribute
    {
        public override FieldType FieldType { get { return FieldType.Number; } }

        public int DecimalPlaces { get; set; }
    }

    public class SPEFCurrencyFieldAttribute : SPEFFieldAttribute
    {
        public override FieldType FieldType { get { return FieldType.Currency; } }

        public int DecimalPlaces { get; set; }
    }

    public class SPEFBooleanFieldAttribute : SPEFFieldAttribute
    {
        public override FieldType FieldType { get { return FieldType.Boolean; } }
    }

    public class SPEFTextFieldAttribute : SPEFFieldAttribute
    {
        public SPEFTextFieldAttribute()
        {
            MaxLength = 255;
            RichText = false;
            RichTextMode = "Compatible";
        }
        public override FieldType FieldType
        {
            get
            {
                if (MaxLength > 255)
                    return FieldType.Note;
                return FieldType.Text;
            }
        }
        public int MaxLength { get; set; }
        public bool RichText { get; set; }
        public string RichTextMode { get; set; }
    }

    public class SPEFDateTimeFieldAttribute : SPEFFieldAttribute
    {
        public SPEFDateTimeFieldAttribute()
        {
            DateTimeFormatType = DateTimeFieldFormatType.DateOnly;
        }
        public override FieldType FieldType { get { return FieldType.DateTime; } }
        /// <summary>
        /// Default: DateOnly
        /// </summary>
        public DateTimeFieldFormatType DateTimeFormatType { get; set; }
    }

    public class SPEFChoiceFieldAttribute : SPEFFieldAttribute
    {
        public SPEFChoiceFieldAttribute()
        {
            Multiple = false;
        }
        public override FieldType FieldType { get { return Multiple ? FieldType.MultiChoice : FieldType.Choice; } }
        public bool Multiple { get; set; }
        public string[] Choices { get; set; }
    }

    public class SPEFLookupFieldAttribute : SPEFFieldAttribute
    {
        public SPEFLookupFieldAttribute()
        {
            List = string.Empty;
            Field = string.Empty;
            Multiple = null;
        }
        public override FieldType FieldType { get { return FieldType.Lookup; } }
        /// <summary>
        /// Viene ignorato e calcolato dal tipo della proprietà. Considerare in future implementazioni di lookup a liste non SPEF
        /// </summary>
        public string List { get; set; }
        /// <summary>
        /// Se non specificato, viene preso Title.
        /// </summary>
        public string Field { get; set; }
        /// <summary>
        /// Se non specificato, viene calcolato dal tipo della proprietà
        /// </summary>
        public bool? Multiple { get; set; }
    }

    public class SPEFUrlFieldAttribute : SPEFFieldAttribute
    {
        public override FieldType FieldType { get { return FieldType.URL; } }
        public UrlFieldFormatType UrlFormatType { get; set; }
    }

    public class SPEFUserFieldAttribute : SPEFFieldAttribute
    {
        public SPEFUserFieldAttribute()
        {
            Multiple = false;
        }
        public override FieldType FieldType { get { return FieldType.User; } }
        public FieldUserSelectionMode UserSelectionMode { get; set; }
        public bool Multiple { get; set; }
    }

    public class SPEFTaxonomyFieldAttribute : SPEFFieldAttribute
    {
        public bool Multiple { get; set; }
        public override FieldType FieldType { get { return FieldType.Lookup; } }
        public string TaxFieldName { get; set; }
        public string TermSetName { get; set; }
        public string TermGroupName { get; set; }
        public string TermStoreName { get; set; }
        public bool TermUserCreated { get; set; }
    }
}
