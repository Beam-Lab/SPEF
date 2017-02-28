using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF.Extensions
{
    public static class ListItemExtensions
    {
        public static string GetStringValue(this ListItem item, string internalName)
        {
            if (item != null)
                return item[internalName] != null ? item[internalName].ToString() : string.Empty;
            else
                return null;
        }

        public static int GetIntValue(this ListItem item, string internalName)
        {
            if (item != null)
            {
                int tempVal = 0;
                return item[internalName] != null && int.TryParse(item[internalName].ToString(), out tempVal)
                ? tempVal : 0;
            }
            else
                return 0;
        }

        public static decimal GetDecimalValue(this ListItem item, string internalName)
        {
            if (item != null)
            {
                decimal tempVal = 0;
                return item[internalName] != null && decimal.TryParse(item[internalName].ToString(), out tempVal)
                ? tempVal : 0;
            }
            else
                return 0;
        }

        public static string[] GetChoicesValues(this ListItem item, string internalName)
        {
            if (item != null)
                return item[internalName] != null ? (string[])item[internalName] : new string[0];
            else
                return null;
        }

        public static string GetLookupValue(this ListItem item, string internalName)
        {
            string value = string.Empty;
            var field = item[internalName] as FieldLookupValue;
            if (field != null)
            {
                return field.LookupValue;
            }
            return value;
        }

        public static int GetLookupIdValue(this ListItem item, string internalName)
        {
            int value = 0;
            var field = item[internalName] as FieldLookupValue;
            if (field != null)
            {
                return field.LookupId;
            }
            return value;
        }

        public static List<string> GetMultiLookupValues(this ListItem item, string internalName)
        {
            var field = item[internalName] as FieldLookupValue[];
            if (field != null)
            {
                return field.Select(f => f.LookupValue).ToList();
            }
            return new List<string>();
        }

        public static List<int> GetMultiLookupIdValues(this ListItem item, string internalName)
        {
            var field = item[internalName] as FieldLookupValue[];
            if (field != null)
            {
                return field.Select(f => f.LookupId).ToList();
            }
            return new List<int>();
        }

        public static int GetUserIdValue(this ListItem item, string internalName)
        {
            int value = 0;
            var field = item[internalName] as FieldUserValue;
            if (field != null)
            {
                return field.LookupId;
            }
            return value;
        }

        public static DateTime GetDateTimeValue(this ListItem item, string internalName)
        {
            DateTime date = DateTime.MinValue;
            DateTime result = DateTime.MinValue;
            if (item != null)
            {
                if (item[internalName] != null && DateTime.TryParse(item[internalName].ToString(), out date))
                    result = date.ToLocalTime();
            }
            return result;
        }

        public static bool GetBoolValue(this ListItem item, string internalName)
        {
            if (item != null)
                return item[internalName] != null ? Convert.ToBoolean(item[internalName].ToString()) : false;
            else
                return false;
        }

        public static string GetUrlValue(this ListItem item, string internalName)
        {
            string value = string.Empty;
            var field = item[internalName] as FieldUrlValue;
            if (field != null)
            {
                return field.Url;
            }
            return value;
        }

        public static KeyValuePair<string, string> GetTaxonomyValue(this ListItem item, string internalName)
        {
            var field = item[internalName] as TaxonomyFieldValue;
            if (field != null)
            {
                var term = new KeyValuePair<string, string>(field.TermGuid, field.Label);
                return term;
            }
            return new KeyValuePair<string, string>();
        }

        public static List<KeyValuePair<string, string>> GetMultiTaxonomyValues(this ListItem item, string internalName)
        {
            var retValues = new List<KeyValuePair<string, string>>();

            var mdColVal = item[internalName] as Dictionary<string, object>;
            if (mdColVal != null)
            {
                var taxValues = mdColVal["_Child_Items_"] as object[];
                foreach (var taxValue in taxValues)
                {
                    var taxDict = taxValue as Dictionary<string, object>;
                    retValues.Add(new KeyValuePair<string, string>(taxDict["TermGuid"].ToString(), taxDict["Label"].ToString()));
                }
            }
            else
            {
                var mdColValTF = item[internalName];
                TaxonomyFieldValueCollection tfvc = mdColValTF as TaxonomyFieldValueCollection;

                if (tfvc != null)
                {
                    foreach (var taxonomyCat in tfvc)
                    {
                        retValues.Add(new KeyValuePair<string, string>(taxonomyCat.TermGuid, taxonomyCat.Label));
                    }
                }
            }
            return retValues;
        }

        public static KeyValuePair<int, string> GetUserValue(this ListItem item, string internalName)
        {
            var field = item[internalName] as FieldUserValue;

            if (field != null)
            {
                return new KeyValuePair<int, string>(field.LookupId, field.LookupValue);
            }
            return new KeyValuePair<int, string>(-1, string.Empty);
        }

        public static List<KeyValuePair<int, string>> GetMultiUserValue(this ListItem item, string internalName)
        {
            var fieldValues = item[internalName] as FieldUserValue[];
            var retList = new List<KeyValuePair<int, string>>();
            if (fieldValues != null)
            {
                foreach(var userValue in fieldValues)
                    retList.Add(new KeyValuePair<int, string>(userValue.LookupId, userValue.LookupValue));
                return retList;
            }
            return new List<KeyValuePair<int, string>>();
        }

        public static int GetUserIDValue(this ListItem item, string internalName)
        {
            var field = item[internalName] as FieldUserValue;

            if (field != null)
            {
                return field.LookupId;
            }
            return -1;
        }

        #region Set

        public static void SetValue(this ListItem item, string internalName, object value)
        {
            item[internalName] = value;
        }

        public static void SetMultiChoiceValue(this ListItem item, string internalName, string[] choices)
        {
            if (choices != null && choices.Length > 0)
                item[internalName] = string.Join(";", choices);
            else
                item[internalName] = string.Empty;
        }
        public static void SetLookupIdValue(this ListItem item, string internalName, int lookupId)
        {
            var field = new FieldLookupValue() { LookupId = lookupId };
            item[internalName] = field;
        }

        public static void SetMultiLookupIdValues(this ListItem item, string internalName, List<int> lookupIds)
        {
            var field = new FieldLookupValue[lookupIds.Count()];
            for (int i = 0; i < lookupIds.Count; i++)
            {
                field[i] = new FieldLookupValue() { LookupId = lookupIds[i] };
            }

            item[internalName] = field;
        }

        public static void SetTaxonomyValue(this ListItem item, string internalName, KeyValuePair<string, string> term)
        {
            string termString = string.Format("-1;#{0}|{1}", term.Value, term.Key);
            item[internalName] = termString;
        }

        public static void SetSPEFTaxonomyValue(this ListItem item, string internalName, SPEF.SPEFTaxonomyItem term)
        {
            string termString = string.Format("-1;#{0}|{1}", term.Value, term.ID);
            item[internalName] = termString;
        }

        public static void SetMultiTaxonomyValues(this ListItem item, string internalName, List<KeyValuePair<string, string>> terms)
        {

            //item["MultiValued"] = "-1;#TermName|00000000-0000-0000-0000-000000000000;#-1;#AnotherTermName|00000000-0000-0000-0000-000000000000";

            var selTerms = terms.Where(o => !(string.IsNullOrWhiteSpace(o.Key))).Select(c => string.Format("-1;#{0}|{1}", c.Value, c.Key));
            var strTerms = string.Join(";#", selTerms);

            item[internalName] = strTerms;
        }

        public static void SetSPEFMultiTaxonomyValues(this ListItem item, string internalName, List<SPEF.SPEFTaxonomyItem> terms)
        {

            //item["MultiValued"] = "-1;#TermName|00000000-0000-0000-0000-000000000000;#-1;#AnotherTermName|00000000-0000-0000-0000-000000000000";

            var selTerms = terms.Where(o => !(o.ID == Guid.Empty)).Select(c => string.Format("-1;#{0}|{1}", c.Value, c.ID));
            var strTerms = string.Join(";#", selTerms);

            item[internalName] = strTerms;
        }

        public static void SetUrlValue(this ListItem item, string internalName, string url)
        {
            var field = new FieldUrlValue() { Url = url };
            item[internalName] = field;
        }

        public static void SetUserValue(this ListItem item, string internalName, KeyValuePair<int, string> user)
        {
            FieldUserValue userValue = new FieldUserValue();
            userValue.LookupId = user.Key;

            item[internalName] = userValue;
        }

        public static void SetUserValue(this ListItem item, string internalName, int userID)
        {
            FieldUserValue userValue = new FieldUserValue();
            userValue.LookupId = userID;

            item[internalName] = userValue;
        }

        public static void SetMultiUserValue(this ListItem item, string internalName, string[] accountNames)
        {
            var usersList = new List<FieldUserValue>();
            foreach (var accountName in accountNames)
            {
                usersList.Add(FieldUserValue.FromUser(accountName));
            }
            item[internalName] = usersList;
        }

        #endregion






    }
}
