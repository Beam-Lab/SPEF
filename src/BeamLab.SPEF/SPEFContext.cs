using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using System.Reflection;
using System.Net;
using System.Linq.Expressions;
using System.Collections;
using Microsoft.SharePoint.Client.Taxonomy;
using BeamLab.SPEF.Models;
using BeamLab.SPEF.Extensions;
using System.Text;
using Microsoft.SharePoint.Client.UserProfiles;
using Newtonsoft.Json;
using BeamLab.SPEF.Constants;

namespace BeamLab.SPEF
{
    public partial class SPEFContext
    {
        #region Properties

        public string ContextUrl { get; set; }
        public string MainVariationLabel { get; set; }

        public ICredentials SPNetworkCredentials { get; set; }

        #endregion

        #region Common

        protected virtual ClientContext GetSharePointContext(string subSite = null)
        {
            return GetSharePointContextWithNetworkCredentials(subSite);
        }

        private ClientContext GetSharePointContextWithNetworkCredentials(string subSite = null)
        {
            try
            {
                var contextUrl = ContextUrl;
                if (!string.IsNullOrWhiteSpace(subSite))
                {
                    if (contextUrl.EndsWith("/"))
                        contextUrl = contextUrl.TrimEnd('/');
                    if (subSite.StartsWith("/"))
                        subSite = subSite.TrimStart('/');
                    contextUrl = $"{contextUrl}/{subSite}";
                }

                var ctx = new ClientContext(contextUrl);
                ctx.Credentials = SPNetworkCredentials;
                return ctx;
            }
            catch (Exception ee)
            {
                throw ee;
            }
        }

        /*
        public string GetLogFolder()
        {
            if (!string.IsNullOrWhiteSpace(logFolder))
                return logFolder;
            var spLogFolder = string.Empty;
            try
            {
                //SPSecurity.RunWithElevatedPrivileges(delegate ()
                //{
                //    var spDiagnosticsService = SPDiagnosticsService.Local;
                //    spLogFolder = spDiagnosticsService.LogLocation;         //recuperare cartella log di SP
                //});
            }
            catch (Exception ex)
            {
                spLogFolder = @"G:\TmpLogs";
            }
            if (string.IsNullOrWhiteSpace(spLogFolder))
                spLogFolder = @"G:\TmpLogs";

            logFolder = spLogFolder;

            return logFolder;
        }
        

        public void WriteLog(string message)
        {
            Log.Information(message);
        }
        
            */
        #endregion

        public SPEFContext(string contextUrl) : this(contextUrl, string.Empty)
        {
        }

        public SPEFContext(string contextUrl, string mainVariationLabel)
        {
            MainVariationLabel = mainVariationLabel.Trim('/');
            ContextUrl = contextUrl;
            structsMapping = new Dictionary<string, SPEFStructInfo>();
            initSPEFContext();
        }

        private Dictionary<string, SPEFStructInfo> structsMapping;
        const string STRUCTS_KEY = "{0}";

        private void initSPEFContext()
        {
            structsMapping.Clear();
            foreach (var propertyInfo in GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                var pType = propertyInfo.PropertyType;
                if (pType.IsGenericType && (pType.GetGenericTypeDefinition() == typeof(List<>)))
                {
                    var listType = pType.GenericTypeArguments[0];
                    var process = listType.IsSubclassOf(typeof(SPEFListItem));

                    if (process)
                    {
                        var listInfo = processStruct(listType);
                    }
                }
            }
        }

        private SPEFStructInfo processStruct(Type itemType)
        {
            var structKey = string.Format(STRUCTS_KEY, itemType.Name);
            if (structsMapping.ContainsKey(structKey))
                return structsMapping[structKey];

            var retStructInfo = new SPEFStructInfo();
            var listAttribute = itemType.GetCustomAttribute<SPEFListAttribute>(false);

            if (listAttribute != null)
            {
                retStructInfo.IsList = true;

                var listTitle = string.IsNullOrWhiteSpace(listAttribute.Title) ? itemType.Name : listAttribute.Title;
                var listDescription = string.IsNullOrWhiteSpace(listAttribute.Description) ? listTitle : listAttribute.Description;
                var listTemplateType = listAttribute.TemplateType;

                retStructInfo.StructType = itemType;
                retStructInfo.Title = listTitle;
                retStructInfo.Description = listDescription;
                retStructInfo.TemplateType = listTemplateType;
                retStructInfo.SubSiteLabel = listAttribute.SubSiteLabel.Trim('/');
                retStructInfo.UseVariations = listAttribute.UseVariations;
                retStructInfo.MainVariationLabel = this.MainVariationLabel;
            }

            var contentTypeAttribute = itemType.GetCustomAttribute<SPEFContentTypeAttribute>(false);
            if (contentTypeAttribute != null)
            {
                retStructInfo.IsContentType = true;

                var contentTypeName = string.IsNullOrWhiteSpace(contentTypeAttribute.Name) ? itemType.Name : contentTypeAttribute.Name;
                var contentTypeDescription = string.IsNullOrWhiteSpace(contentTypeAttribute.Description) ? contentTypeName : contentTypeAttribute.Description;

                retStructInfo.StructType = itemType;
                retStructInfo.ContentTypeName = contentTypeName;
                retStructInfo.Description = contentTypeDescription;
            }

            var baseType = itemType.BaseType;
            while (baseType != null)
            {
                var ctAttribute = baseType.GetCustomAttribute<SPEFContentTypeAttribute>(false);
                var setCT = retStructInfo.IsList && string.IsNullOrWhiteSpace(retStructInfo.ContentTypeName);
                var setBaseCT = retStructInfo.IsContentType && string.IsNullOrWhiteSpace(retStructInfo.BaseContentTypeName);
                if (ctAttribute != null)
                {
                    var contentTypeName = string.IsNullOrWhiteSpace(ctAttribute.Name) ? baseType.Name : ctAttribute.Name;
                    if (setBaseCT)
                        retStructInfo.BaseContentTypeName = contentTypeName;
                    if (setCT)
                        retStructInfo.ContentTypeName = contentTypeName;
                }
                var lsAttribute = baseType.GetCustomAttribute<SPEFListAttribute>(false);
                if (lsAttribute != null || ctAttribute != null)
                {
                    processStruct(baseType);
                    break;
                }
                baseType = baseType.BaseType;
            }

            var itemProperties = itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var itemProperty in itemProperties)
            {
                var fieldInfo = processProperty(itemProperty);

                if (fieldInfo != null)
                {
                    retStructInfo.FieldsInfo.Add(fieldInfo);
                }
            }

            structsMapping.Add(structKey, retStructInfo);

            return retStructInfo;
        }

        public List<string> Init()
        {
            foreach (var contentTypeInfo in structsMapping.Values.Where(s => s.IsContentType))
            {
                initContentType(contentTypeInfo);
            }

            var retMessages = new List<string>();
            try
            {
                foreach (var listInfo in structsMapping.Values.Where(s => s.IsList))
                {
                    if (initList(listInfo))
                        retMessages.Add(string.Format("List OK: {0}", listInfo.Title));
                }
            }
            catch (Exception ex)
            {
                retMessages.Add(string.Format("ERROR: {0}", ex.Message));
                return retMessages;
            }
            return retMessages;
        }

        public Type Resolve(string title)
        {
            var lowerTitle = title.ToLower();
            var structMapping = structsMapping.Values.Where(s => s.Title.ToLower() == lowerTitle).FirstOrDefault();

            if (structMapping != null)
                return structMapping.StructType;

            return null;
        }

        const string contentTypeGroup = "MY GROUP";
        private bool initContentType(SPEFStructInfo contentTypeInfo)
        {
            using (var context = GetSharePointContext())
            {
                var web = context.Web;

                var contentTypeCollection = web.ContentTypes;
                var contentTypeName = contentTypeInfo.ContentTypeName;
                context.Load(contentTypeCollection, ct => ct.Include(l => l.Name).Where(l => l.Name == contentTypeName));
                context.ExecuteQuery();

                ContentType contentType = null;
                if (contentTypeCollection.Count > 0)
                    contentType = contentTypeCollection[0];
                else
                {
                    ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
                    newCt.Name = contentTypeName;
                    newCt.Group = contentTypeGroup;

                    var baseContentTypeName = string.IsNullOrWhiteSpace(contentTypeInfo.BaseContentTypeName) ? "Item" : contentTypeInfo.BaseContentTypeName;
                    var baseContentTypes = context.LoadQuery(web.ContentTypes.Where(ct => ct.Name == baseContentTypeName));
                    context.ExecuteQuery();
                    var baseContentType = baseContentTypes.FirstOrDefault();

                    if (baseContentType == null)
                    {
                        var newCTID = "0x0100" + Guid.NewGuid().ToString().Replace("-", "");
                        newCt.Id = newCTID;
                    }
                    else
                    {
                        newCt.ParentContentType = baseContentType;
                    }

                    contentType = web.ContentTypes.Add(newCt);
                    context.ExecuteQuery();
                }

                foreach (var fieldInfo in contentTypeInfo.FieldsInfo)
                {
                    var fieldCollection = contentType.Fields;
                    var internalName = fieldInfo.Name;
                    context.Load(fieldCollection, fields => fields.Include(f => f.InternalName).Where(f => f.InternalName == internalName));
                    context.ExecuteQuery();
                    //check if content type contains this field
                    if (fieldCollection.Count > 0)
                    {
                        var field = fieldCollection[0];
                        context.Load(field);
                        context.ExecuteQuery();

                        updateField(field, fieldInfo);

                        field.Update();
                        context.ExecuteQuery();
                    }
                    else
                    {
                        //check if site field exists
                        Field field = null;
                        fieldCollection = web.Fields;
                        context.Load(fieldCollection, fields => fields.Include(f => f.StaticName).Where(f => f.StaticName == internalName));
                        context.ExecuteQuery();
                        if (fieldCollection.Count > 0)
                        {
                            field = fieldCollection[0];
                        }
                        else
                        {
                            //create site field
                            if (fieldInfo.Type == FieldType.Lookup && fieldInfo.LookupListID == Guid.Empty)
                            {
                                var targetList = web.Lists.GetByTitle(fieldInfo.LookupList);
                                context.Load(targetList);
                                context.ExecuteQuery();
                                if (targetList == null)
                                    continue;
                                fieldInfo.LookupListID = targetList.Id;
                            }

                            var fieldXml = fieldInfo.GetXml();
                            field = fieldCollection.AddFieldAsXml(fieldXml, true, AddFieldOptions.AddFieldInternalNameHint);
                            context.ExecuteQuery();
                        }
                        //add site field to contentType
                        var fieldLinkCreationInformation = new FieldLinkCreationInformation();
                        fieldLinkCreationInformation.Field = field;
                        contentType.FieldLinks.Add(fieldLinkCreationInformation);
                        contentType.Update(true);
                        context.ExecuteQuery();
                    }
                }
            }
            return true;
        }

        private bool initList(SPEFStructInfo listInfo, string contextUrl = null)
        {
            using (var context = GetSharePointContext(contextUrl ?? listInfo.ContextUrl))
            {
                var web = context.Web;

                var listCollection = web.Lists;
                var listTitle = listInfo.Title;
                context.Load(listCollection, lists => lists.Include(l => l.Title).Where(l => l.Title == listTitle));
                context.ExecuteQuery();

                List list = null;
                if (listCollection.Count > 0)
                    list = listCollection[0];
                else
                {
                    var newListCreationInformation = new ListCreationInformation();
                    newListCreationInformation.Title = listInfo.Title;
                    newListCreationInformation.Description = listInfo.Description;
                    newListCreationInformation.TemplateType = (int)listInfo.TemplateType;

                    try
                    {
                        list = web.Lists.Add(newListCreationInformation);
                        context.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("Exception in creation list {0}", listInfo.Title), ex);
                    }
                }

                if (!string.IsNullOrWhiteSpace(listInfo.ContentTypeName))
                {
                    var listContentTypeCollection = list.ContentTypes;
                    var contentTypeName = listInfo.ContentTypeName;
                    context.Load(listContentTypeCollection, ct => ct.Include(l => l.Name).Where(l => l.Name == contentTypeName));
                    context.ExecuteQuery();
                    //if (listContentTypeCollection.Where(c => c.Name == contentTypeName).FirstOrDefault() == null)
                    if (listContentTypeCollection.Count == 0)
                    {
                        try
                        {
                            list.ContentTypesEnabled = true;
                            list.Update();

                            using (var mainContext = GetSharePointContext())
                            {
                                var webContentTypeCollection = mainContext.Web.ContentTypes;
                                mainContext.Load(webContentTypeCollection, ct => ct.Include(l => l.Name).Where(l => l.Name == contentTypeName));
                                mainContext.ExecuteQuery();

                                ContentType contentType = null;
                                if (webContentTypeCollection.Count > 0)
                                {
                                    contentType = webContentTypeCollection[0];
                                    var listWeb = web;
                                    if (!string.IsNullOrWhiteSpace(listInfo.ContextUrl))
                                        listWeb = mainContext.Site.OpenWeb(listInfo.ContextUrl);

                                    var ctListCollection = listWeb.Lists;
                                    mainContext.Load(ctListCollection, lists => lists.Include(l => l.Title).Where(l => l.Title == listTitle));
                                    mainContext.ExecuteQuery();

                                    List ctList = null;
                                    if (ctListCollection.Count > 0)
                                    {
                                        ctList = ctListCollection[0];

                                        ctList.ContentTypes.AddExistingContentType(contentType);
                                        ctList.Update();
                                        mainContext.ExecuteQuery();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(string.Format("Exception adding content type {0} to list {1}", contentTypeName, listInfo.Title), ex);
                        }
                    }
                }

                foreach (var fieldInfo in listInfo.FieldsInfo)
                {
                    if (fieldInfo.Ignore)
                        continue;

                    var fieldCollection = list.Fields;
                    var internalName = fieldInfo.Name;
                    context.Load(fieldCollection, fields => fields.Include(f => f.InternalName).Where(f => f.InternalName == internalName));
                    context.ExecuteQuery();

                    if (fieldCollection.Count > 0)
                    {
                        var field = fieldCollection[0];
                        context.Load(field);
                        context.ExecuteQuery();

                        try
                        {
                            updateField(field, fieldInfo);
                            field.Update();
                            context.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(string.Format("Exception updating field {0} in list {1}", fieldInfo.Name, listInfo.Title), ex);
                        }
                    }
                    else
                    {
                        try
                        {
                            if (fieldInfo.Type == FieldType.Lookup && !fieldInfo.FieldTypeTaxonomy && fieldInfo.LookupListID == Guid.Empty)
                            {
                                var targetList = web.Lists.GetByTitle(fieldInfo.LookupList);
                                context.Load(targetList);
                                context.ExecuteQuery();
                                if (targetList == null)
                                    continue;
                                fieldInfo.LookupListID = targetList.Id;
                            }
                            if (fieldInfo.Type == FieldType.Lookup && fieldInfo.FieldTypeTaxonomy)
                            {
                                //  https://andrewwburns.com/2013/12/18/working-with-the-taxonomy-in-csom/
                                var taxNoteFieldXml = string.Format(@"<Field Type='Note' DisplayName='__{0}_0' StaticName='__{0}TaxHTField0' Name='__{0}TaxHTField0' 
                                    ShowInViewForms='FALSE' Required = 'FALSE' Hidden='TRUE' CanToggleHidden='TRUE' ShowField='Term1033'
                                    Overwrite ='TRUE' /> ", fieldInfo.XmlName);
                                var taxNoteField = list.Fields.AddFieldAsXml(taxNoteFieldXml, true, AddFieldOptions.AddFieldInternalNameHint);
                                context.Load(taxNoteField);
                                context.ExecuteQuery();

                                Guid taxNoteFieldId = taxNoteField.Id;
                                Guid termStoreGuid = Guid.Empty;
                                Guid termGroupGuid = Guid.Empty;
                                Guid termSetGuid = Guid.Empty;

                                getTaxonomyServiceGuid(fieldInfo.TermStoreName, fieldInfo.TermGroupName, fieldInfo.TermSetName,
                                    out termStoreGuid, out termGroupGuid, out termSetGuid);

                                fieldInfo.TaxFieldId = taxNoteFieldId;
                                fieldInfo.TermStoreId = termGroupGuid;
                                fieldInfo.TermGroupId = termGroupGuid;
                                fieldInfo.TermSetId = termSetGuid;
                            }

                            var fieldXml = fieldInfo.GetXml();
                            var field = list.Fields.AddFieldAsXml(fieldXml, true, AddFieldOptions.AddFieldInternalNameHint);
                            context.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(string.Format("Exception creating field {0} in list {1}", fieldInfo.Name, listInfo.Title), ex);
                        }
                    }
                }
            }
            return true;
        }

        private bool updateField(Field field, SPEFFieldInfo fieldInfo)
        {
            //DisplayName
            if (!string.IsNullOrWhiteSpace(fieldInfo.Title) && field.Title != fieldInfo.Title)
            {
                field.Title = fieldInfo.Title;
                //SetFieldDisplayName(field, pField.Title);
            }
            field.Required = fieldInfo.Required;
            //if (field.Hidden != fieldInfo.Hidden)
            //    field.Hidden = fieldInfo.Hidden;
            //if (field.sh != pField.ShowInDisplayForm)
            if (fieldInfo.Type != FieldType.Choice && fieldInfo.Type != FieldType.MultiChoice)
            {
                field.SetShowInNewForm(fieldInfo.ShowInNewForm);
                field.SetShowInDisplayForm(fieldInfo.ShowInDisplayForm);
                field.SetShowInEditForm(fieldInfo.ShowInEditForm);
            }
            if (fieldInfo.Type != FieldType.Invalid && field.FieldTypeKind != fieldInfo.Type && !fieldInfo.FieldTypeTaxonomy)
                field.FieldTypeKind = fieldInfo.Type;

            if (fieldInfo.Type == FieldType.Text)
            {
                var textFieldAttribute = field as FieldText;
                textFieldAttribute.MaxLength = fieldInfo.MaxLength;
                //textFieldAttribute.RichText = fieldInfo.RichText;
                //textFieldAttribute.RichTextMode = fieldInfo.RichTextMode;
            }

            if (fieldInfo.Type == FieldType.DateTime)
            {
                var dateTimeFieldAttribute = field as FieldDateTime;
                if (dateTimeFieldAttribute != null && dateTimeFieldAttribute.DisplayFormat != fieldInfo.DateTimeFormatType)
                    dateTimeFieldAttribute.DisplayFormat = fieldInfo.DateTimeFormatType;
            }

            if (fieldInfo.Type == FieldType.Choice || fieldInfo.Type == FieldType.MultiChoice)
            {
                var cf = field as FieldMultiChoice;
                if (cf != null)
                {
                    var updateChoices = false;
                    var fieldChoices = cf.Choices.ToList();
                    for (int i = 0; i < fieldInfo.Choices.Count; i++)
                    {
                        if (!fieldChoices.Contains(fieldInfo.Choices[i]))
                        {
                            updateChoices = true;
                            fieldChoices.Add(fieldInfo.Choices[i]);
                        }
                    }
                    if (updateChoices)
                        cf.Choices = fieldChoices.ToArray();
                }
            }

            if (fieldInfo.Type == FieldType.URL)
            {
                var uf = field as FieldUrl;
                //if(uf != null)
            }

            if (fieldInfo.Type == FieldType.User)
            {
                var uf = field as FieldUser;
                if (uf != null && uf.SelectionMode != fieldInfo.UserSelectionMode)
                {
                    uf.SelectionMode = fieldInfo.UserSelectionMode;
                }
            }

            foreach (var attribute in fieldInfo.AdditionalAttributes)
            {
                var propInfo = field.GetType().GetProperty(attribute.Key);
                if (propInfo == null)
                    continue;
                object propertyValue = propInfo.GetValue(field, null);

                if (propertyValue == null || propertyValue.ToString() != attribute.Value)
                {
                    try
                    {
                        var newValue = Convert.ChangeType(attribute.Value, propInfo.PropertyType);
                        propInfo.SetValue(field, newValue);
                    }
                    catch (Exception ex)
                    {
                        propInfo.SetValue(field, propertyValue);
                    }
                }
            }

            return true;
        }

        private SPEFFieldInfo processProperty(PropertyInfo propertyInfo)
        {
            var retspefFieldInfo = new SPEFFieldInfo(propertyInfo);

            var fieldAttribute = propertyInfo.GetCustomAttribute<SPEFFieldAttribute>();

            if (fieldAttribute == null)
            {
                //creare standard in base al tipo della proprietà
                return null;
            }

            retspefFieldInfo.Name = (string.IsNullOrWhiteSpace(fieldAttribute.Name)) ? propertyInfo.Name : fieldAttribute.Name;
            retspefFieldInfo.Title = (string.IsNullOrWhiteSpace(fieldAttribute.Title)) ? propertyInfo.Name : fieldAttribute.Title;

            retspefFieldInfo.Ignore = fieldAttribute.Ignore;
            retspefFieldInfo.Readonly = fieldAttribute.Readonly;
            retspefFieldInfo.Required = fieldAttribute.Required;
            retspefFieldInfo.Type = fieldAttribute.FieldType;
            retspefFieldInfo.ShowInNewForm = fieldAttribute.ShowInNew;
            retspefFieldInfo.ShowInDisplayForm = fieldAttribute.ShowInDisplay;
            retspefFieldInfo.ShowInEditForm = fieldAttribute.ShowInEdit;
            retspefFieldInfo.ShowInViewForm = fieldAttribute.ShowInView;

            if (fieldAttribute is SPEFNumericFieldAttribute)
            {
                var numericFieldAttribute = fieldAttribute as SPEFNumericFieldAttribute;
                retspefFieldInfo.DecimalPlaces = numericFieldAttribute.DecimalPlaces;
            }
            if (fieldAttribute is SPEFCurrencyFieldAttribute)
            {
                var currencyFieldAttribute = fieldAttribute as SPEFCurrencyFieldAttribute;
                if (currencyFieldAttribute.DecimalPlaces > 0)
                    retspefFieldInfo.DecimalPlaces = Math.Min(currencyFieldAttribute.DecimalPlaces, 5);
            }
            else if (fieldAttribute is SPEFBooleanFieldAttribute)
            {
            }
            else if (fieldAttribute is SPEFTextFieldAttribute)
            {
                var textFieldAttribute = fieldAttribute as SPEFTextFieldAttribute;
                retspefFieldInfo.MaxLength = textFieldAttribute.MaxLength;
                retspefFieldInfo.RichText = textFieldAttribute.RichText;
                retspefFieldInfo.RichTextMode = textFieldAttribute.RichTextMode;
            }
            else if (fieldAttribute is SPEFChoiceFieldAttribute)
            {
                var choiceFieldAttribute = fieldAttribute as SPEFChoiceFieldAttribute;
                retspefFieldInfo.Multiple = choiceFieldAttribute.Multiple;
                retspefFieldInfo.Choices = choiceFieldAttribute.Choices.ToList();
            }
            else if (fieldAttribute is SPEFLookupFieldAttribute)
            {
                SPEFStructInfo structInfo = null;
                var type = propertyInfo.PropertyType;
                var isMultiple = false;
                if (type.IsGenericType)
                {
                    var argTypes = type.GetGenericArguments();
                    var t1 = argTypes[0];
                    isMultiple = true;
                    structInfo = getSPEFStructureInfo(t1);
                }
                else
                    structInfo = getSPEFStructureInfo(type);

                var lookupFieldAttribute = fieldAttribute as SPEFLookupFieldAttribute;
                retspefFieldInfo.Multiple = lookupFieldAttribute.Multiple ?? isMultiple;
                retspefFieldInfo.LookupList = structInfo.Title;
                retspefFieldInfo.LookupField = string.IsNullOrWhiteSpace(lookupFieldAttribute.Field) ? StandardListItemFields.Title : lookupFieldAttribute.Field;
            }
            else if (fieldAttribute is SPEFTaxonomyFieldAttribute)
            {
                var taxonomyFieldAttribute = fieldAttribute as SPEFTaxonomyFieldAttribute;
                retspefFieldInfo.Multiple = taxonomyFieldAttribute.Multiple;
                retspefFieldInfo.FieldTypeTaxonomy = true;
                retspefFieldInfo.TermStoreName = taxonomyFieldAttribute.TermStoreName;
                retspefFieldInfo.TermGroupName = taxonomyFieldAttribute.TermGroupName;
                retspefFieldInfo.TermSetName = taxonomyFieldAttribute.TermSetName;
                retspefFieldInfo.TermUserCreated = taxonomyFieldAttribute.TermUserCreated;
            }
            else if (fieldAttribute is SPEFUrlFieldAttribute)
            {
                var urlFieldAttribute = fieldAttribute as SPEFUrlFieldAttribute;
                retspefFieldInfo.UrlFormatType = urlFieldAttribute.UrlFormatType;
            }
            else if (fieldAttribute is SPEFUserFieldAttribute)
            {
                var userFieldAttribute = fieldAttribute as SPEFUserFieldAttribute;
                retspefFieldInfo.UserSelectionMode = userFieldAttribute.UserSelectionMode;
                retspefFieldInfo.Multiple = userFieldAttribute.Multiple;
            }

            return retspefFieldInfo;
        }

        #region util

        private SPEFStructInfo getSPEFStructureInfo(Type t)
        {
            try
            {
                var typeName = t.Name;
                var currKey = string.Format(STRUCTS_KEY, typeName);
                var currListInfo = structsMapping[currKey];

                return currListInfo;
            }
            catch (Exception ex)
            {
                throw new Exception($"{ex.Message} - {t.Name}");
            }
        }

        private SPEFStructInfo getSPEFStructureInfo<T>()
        {
            var typeName = typeof(T).Name;
            var currKey = string.Format(STRUCTS_KEY, typeName);
            var currListInfo = structsMapping[currKey];

            return currListInfo;
        }
        #endregion

        #region util Taxonomy

        private TermSet getTaxonomyServiceGuid(string termStoreName, string termGroupName, string termSetName,
                out Guid termStoreGuid, out Guid termGroupGuid, out Guid termSetGuid)
        {
            termStoreGuid = Guid.Empty;
            termSetGuid = Guid.Empty;
            termGroupGuid = Guid.Empty;

            using (var clientContext = GetSharePointContext())
            {
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
                if (taxonomySession == null)
                {
                    return null;
                }

                clientContext.Load(taxonomySession);
                clientContext.ExecuteQuery();
                clientContext.Load(taxonomySession.TermStores);
                clientContext.ExecuteQuery();
                if (taxonomySession.TermStores == null)
                {
                    return null;
                }

                TermStore termStore = null;
                var stores = string.Empty;

                termStore = taxonomySession.TermStores.GetByName(termStoreName);
                if (termStore == null)
                    return null;
                clientContext.Load(termStore,
                    store => store.Name,
                    store => store.Id,
                    store => store.Groups.Include(
                        group => group.Name,
                        group => group.Id,
                        group => group.TermSets.Include(
                            set => set.Name,
                            set => set.Id
                        )
                    )
                );
                clientContext.ExecuteQuery();
                termStoreGuid = termStore.Id;

                TermGroup termGroup = null;
                foreach (var group in termStore.Groups)
                {
                    if (group.Name == termGroupName)
                    {
                        termGroup = group;
                        break;
                    }
                }

                if (termGroup == null)
                    return null;
                termGroupGuid = termGroup.Id;

                TermSet termSet = null;
                foreach (var set in termGroup.TermSets)
                {
                    if (set.Name == termSetName)
                    {
                        termSet = set;
                        break;
                    }
                }
                if (termSet == null)
                    return null;
                termSetGuid = termSet.Id;

                return termSet;
            }
        }

        #endregion

        #region Count

        public List<SPEFCount<TP>> LoadCount<T, TP>(Expression<Func<T, TP>> fieldExpression, string label = null) where T : SPEFListItem
        {
            var retList = new List<SPEFCount<TP>>();
            var currListInfo = getSPEFStructureInfo<T>();
            var fieldInfo = getPropertyInfo(fieldExpression);
            var listFieldInfo = currListInfo.FieldsInfo.Where(f => f.PropertyInfo.Name == fieldInfo.Name).FirstOrDefault();
            var spQuery = $"<View><ViewFields><FieldRef Name=\"{listFieldInfo.Name}\" /></ViewFields></View>";

            using (var context = GetSharePointContext(label ?? currListInfo.ContextUrl))
            {
                try
                {
                    var web = context.Web;
                    var listCollection = web.Lists;
                    var listTitle = currListInfo.Title;
                    context.Load(listCollection, lists => lists.Include(l => l.Title).Where(l => l.Title == listTitle));
                    context.ExecuteQuery();

                    if (listCollection.Count == 0)
                        return null;

                    var list = listCollection[0];

                    var camlQuery = new CamlQuery() { ViewXml = spQuery };

                    var listItemCollection = list.GetItems(camlQuery);
                    context.Load(listItemCollection);

                    context.ExecuteQuery();

                    var dictRes = new Dictionary<string, SPEFCount<TP>>();

                    foreach (var countItem in listItemCollection)
                    {
                        var fieldValue = getFieldValue(countItem, listFieldInfo);
                        var newValue = (TP)fieldValue;
                        var dictKey = JsonConvert.SerializeObject(newValue);
                        if (!dictRes.ContainsKey(dictKey))
                        {
                            var tmpCountItem = new SPEFCount<TP>()
                            {
                                Value = newValue,
                                Count = 0
                            };
                            dictRes.Add(dictKey, tmpCountItem);
                        }
                        dictRes[dictKey].Count++;
                    }
                    retList = dictRes.Values.ToList();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Exception loading items from query ( {spQuery} ). list {currListInfo.Title}", ex);
                }
            }
            return retList;
        }

        #endregion

        #region Retrieve Objects

        public object LoadByID(Type t, int id, string label = null)
        {
            if (id <= 0)
            {
                return null;
            }

            var mi = GetType().GetMethod("LoadByID");
            var fooRef = mi.MakeGenericMethod(t);
            var retObj = fooRef.Invoke(this, new object[] { id, label });

            return retObj;
        }

        public object TryLoad(Type t, string label = null)
        {
            var mi = GetType().GetMethod("LoadItems");
            var fooRef = mi.MakeGenericMethod(t);
            var retObj = fooRef.Invoke(this, new object[] { label });

            return retObj;
        }

        public bool Load<T>(int id, out T entity, string label = null) where T : SPEFListItem
        {
            if (id <= 0)
            {
                entity = null;
                return false;
            }

            var retObj = LoadByID<T>(id, label);
            entity = retObj;

            var retVal = retObj != null;
            return retVal;
        }
        public T LoadByID<T>(int id, string label = null) where T : SPEFListItem
        {
            var currListInfo = getSPEFStructureInfo<T>();

            using (var context = GetSharePointContext(label ?? currListInfo.ContextUrl))
            {
                var web = context.Web;
                var listCollection = web.Lists;
                var listTitle = currListInfo.Title;
                context.Load(listCollection, lists => lists.Include(l => l.Title, l => l.EnableAttachments).Where(l => l.Title == listTitle));
                context.ExecuteQuery();

                List list = null;
                if (listCollection.Count == 0)
                    return null;
                list = listCollection[0];

                var currListItem = list.GetItemById(id);
                context.Load(currListItem);
                if (list.EnableAttachments)
                    context.Load(currListItem.AttachmentFiles);
                try
                {
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Exception loading item ID {id} from list {listTitle}", ex);
                }
                var retObj = loadInto<T>(currListItem);
                return retObj;
            }
        }


        private object getFieldValue(ListItem listItem, SPEFFieldInfo fieldInfo)
        {
            switch (fieldInfo.Type)
            {
                case FieldType.Note:
                case FieldType.Text:
                    return listItem.GetStringValue(fieldInfo.Name);
                case FieldType.Boolean:
                    return listItem.GetBoolValue(fieldInfo.Name);
                case FieldType.Number:
                    if (fieldInfo.DecimalPlaces > 0)
                        return listItem.GetDecimalValue(fieldInfo.Name);
                    return listItem.GetIntValue(fieldInfo.Name);
                case FieldType.Currency:
                    return listItem.GetDecimalValue(fieldInfo.Name);
                case FieldType.DateTime:
                    return listItem.GetDateTimeValue(fieldInfo.Name);
                case FieldType.Choice:
                    return listItem.GetStringValue(fieldInfo.Name);
                case FieldType.MultiChoice:
                    {
                        var strVal = listItem.GetChoicesValues(fieldInfo.Name);
                        return strVal;
                    }
                case FieldType.Lookup:
                    {
                        if (fieldInfo.FieldTypeTaxonomy)
                        {
                            if (fieldInfo.Multiple)
                            {
                                var retList = listItem.GetMultiTaxonomyValues(fieldInfo.Name);
                                return retList.Select(t => new SPEFTaxonomyItem() { ID = Guid.Parse(t.Key), Value = t.Value }).ToList();
                            }
                            else
                            {
                                var retTerm = listItem.GetTaxonomyValue(fieldInfo.Name);
                                return new SPEFTaxonomyItem() { ID = Guid.Parse(retTerm.Key), Value = retTerm.Value };
                            }
                        }
                        else
                        {
                            var refObjType = fieldInfo.PropertyInfo.PropertyType;
                            var refObj = Activator.CreateInstance(refObjType);
                            if (!fieldInfo.Multiple)
                            {
                                var refObjItem = refObj as SPEFItem;
                                refObjItem.ID = listItem.GetLookupIdValue(fieldInfo.Name);
                                refObjItem.DisplayField = listItem.GetLookupValue(fieldInfo.Name);
                            }
                            else
                            {
                                var values = listItem.GetMultiLookupValues(fieldInfo.Name);
                                var ids = listItem.GetMultiLookupIdValues(fieldInfo.Name);
                                var tmpList = new SPEFItem[ids.Count];
                                for (int i = 0; i < ids.Count; i++)
                                {
                                    var lookupObjType = Activator.CreateInstance(refObjType.GenericTypeArguments[0]);
                                    var lookupObj = lookupObjType as SPEFItem;
                                    lookupObj.ID = ids[i];
                                    lookupObj.DisplayField = values[i];

                                    refObjType.GetMethod("Add").Invoke(refObj, new[] { lookupObjType });
                                }

                            }
                            return refObj;
                        }
                    }
                case FieldType.URL:
                    {
                        return listItem.GetUrlValue(fieldInfo.Name);
                    }
                case FieldType.User:
                    if (!fieldInfo.Multiple)
                    {
                        var userPair = listItem.GetUserValue(fieldInfo.Name);
                        return new SPEFUser(userPair.Key)
                        {
                            DisplayName = userPair.Value
                        };
                    }
                    else
                    {
                        var usersPairs = listItem.GetMultiUserValue(fieldInfo.Name);
                        var retList = usersPairs.Select(u => new SPEFUser(u.Key)
                        {
                            DisplayName = u.Value
                        }).ToList();
                        return retList;
                    }
                default:
                    return listItem.GetStringValue(fieldInfo.Name);
            }
        }

        public IEnumerable<T> LoadItems<T>(string label = null) where T : SPEFListItem
        {
            return Load<T>(null, label);
        }

        public IEnumerable<T> Load<T>(string label = null, int rowsLimit = 0) where T : SPEFListItem
        {
            return Load<T>(null, label, rowsLimit);
        }

        public IEnumerable<T> Load<T>(ISPEFQueryNode<T> node, string label = null, int rowsLimit = 0, int lastRowID = 0) where T : SPEFListItem
        {
            var currListInfo = getSPEFStructureInfo<T>();

            if (!string.IsNullOrWhiteSpace(currListInfo.ContentTypeName))
            {
                if (string.IsNullOrWhiteSpace(currListInfo.ContentTypeID))
                {
                    var contentTypeID = getContentTypeID(currListInfo.ContentTypeName);
                    if (contentTypeID != null)
                        currListInfo.ContentTypeID = contentTypeID.StringValue;
                }
            }

            var spQuery = string.Empty;
            SPEFFieldInfo sortListFieldInfo = null;
            if (node != null)
            {
                //if(!string.IsNullOrWhiteSpace(currListInfo.ContentTypeID))
                //    spQuery = spQuery.and
                spQuery = getQuery(node, out sortListFieldInfo);
            }
            var rowsLimitQuery = string.Empty;
            if (rowsLimit > 0)
            {
                rowsLimitQuery = string.Format("<RowLimit Paged='{1}'>{0}</RowLimit>", rowsLimit, (lastRowID > 0));
            }

            //
            var fieldsFilter = new StringBuilder();
            foreach (var fieldInfo in currListInfo.FieldsInfo)
            {
                fieldsFilter.Append(string.Format("<FieldRef Name='{0}'/>", fieldInfo.Name));
            }
            //

            var xmlQuery = string.Empty;

            using (var context = GetSharePointContext(label ?? currListInfo.ContextUrl))
            {
                try
                {
                    var web = context.Web;
                    var listCollection = web.Lists;
                    var listTitle = currListInfo.Title;
                    context.Load(listCollection, lists => lists.Include(l => l.Title, l => l.EnableAttachments).Where(l => l.Title == listTitle));
                    context.ExecuteQuery();

                    if (listCollection.Count == 0)
                        return null;

                    var list = listCollection[0];

                    var attachmentsField = list.EnableAttachments ? "<FieldRef Name=\"Attachments\"></FieldRef>" : string.Empty;
                    xmlQuery = string.Format("<View><ViewFields>{2}{3}</ViewFields><Query>{0}</Query>{1}</View>", spQuery, rowsLimitQuery, fieldsFilter, attachmentsField);

                    var camlQuery = new CamlQuery() { ViewXml = xmlQuery };
                    if (rowsLimit > 0 && lastRowID > 0 && sortListFieldInfo != null)
                    {
                        var lastItem = list.GetItemById(lastRowID);
                        context.Load(lastItem);
                        context.ExecuteQuery();

                        var lastItemValue = lastItem[sortListFieldInfo.Name];

                        camlQuery.ListItemCollectionPosition = new ListItemCollectionPosition() { PagingInfo = string.Format("Paged=TRUE&p_ID={0}&p_{1}={2}", lastRowID, sortListFieldInfo.Name, lastItemValue) };
                    }

                    var listItemCollection = list.GetItems(camlQuery);
                    context.Load(listItemCollection);
                    if (list.EnableAttachments)
                        context.Load(listItemCollection, items => items.Include(i => i.AttachmentFiles));

                    context.ExecuteQuery();
                    return loadInto<T>(listItemCollection);
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Exception loading items from query ( {0} ). list {1}", xmlQuery, currListInfo.Title), ex);
                }
            }
        }

        private string getQuery<T>(ISPEFQueryNode<T> node, out SPEFFieldInfo sortListFieldInfo) where T : SPEFListItem
        {
            sortListFieldInfo = null;
            var retQuery = string.Empty;

            var t = node.GetType();
            var currListInfo = getSPEFStructureInfo<T>();
            if (t.IsGenericType && t.Name == (typeof(SPEFOperation<>)).Name) //t.GetGenericTypeDefinition() == (typeof(SPEFQueryNode<>)))
            {
                var targetType = typeof(SPEFOperation<>).MakeGenericType(typeof(T));
                dynamic nodeFilter = Convert.ChangeType(node, targetType);
                var filterQuery = getFilterQueryNonRec(nodeFilter);
                //var filterTest = getFilterQuery(nodeFilter);

                retQuery = string.Format("<Where>{0}</Where>", filterQuery);
                return retQuery;
            }
            else if (t.IsGenericType && t.Name == (typeof(SPEFExpression<,>)).Name)  // t.GetGenericTypeDefinition() == (typeof(SPEFQueryNode<>)))
            {
                var exprTypes = t.GetGenericArguments();
                var t1 = exprTypes[1];
                var targetType = typeof(SPEFExpression<,>).MakeGenericType(typeof(T), t1);
                dynamic nodeFilter = Convert.ChangeType(node, targetType);
                var filterQuery = getFilterQueryNonRec(nodeFilter);
                retQuery = string.Format("<Where>{0}</Where>", filterQuery);
                return retQuery;
            }
            else if (t.IsGenericType && t.GetGenericTypeDefinition() == (typeof(SPEFSortNode<,>)))
            {
                var exprTypes = t.GetGenericArguments();
                var t1 = exprTypes[1];
                var targetType = typeof(SPEFSortNode<,>).MakeGenericType(typeof(T), t1);

                dynamic nodeSort = Convert.ChangeType(node, targetType);
                retQuery = getSortQuery(nodeSort, out sortListFieldInfo);

                return retQuery;
            }


            return string.Empty;
        }


        private string getFilterQuery<T>(SPEFQueryNode<T> node)
        {
            var t = node.GetType();
            if (t.IsGenericType && t.GetGenericTypeDefinition() == (typeof(SPEFExpression<,>)))
            {
                var exprTypes = t.GetGenericArguments();
                var t1 = exprTypes[1];
                var targetType = typeof(SPEFExpression<,>).MakeGenericType(typeof(T), t1);

                dynamic nodeExpression = Convert.ChangeType(node, targetType);

                var fieldInfo = getPropertyInfo(nodeExpression.Expression);

                var currListInfo = getSPEFStructureInfo<T>();

                var listFieldInfo = currListInfo.FieldsInfo.Where(f => f.PropertyInfo.Name == fieldInfo.Name).FirstOrDefault();
                if (listFieldInfo != null)
                {
                    var op = (Op)nodeExpression.Op;
                    if (!(op == Op.IsNotNull || op == Op.IsNull))
                    {
                        var value = nodeExpression.Value;
                        if (listFieldInfo.Type == FieldType.Lookup && !listFieldInfo.FieldTypeTaxonomy)
                        {
                            var isID = true;
                            var spefObj = nodeExpression.Value as SPEFListItem;
                            if (spefObj != null)
                                value = spefObj.ID;
                            else
                            {
                                if (!(nodeExpression.Value is int))
                                    isID = false;
                                value = nodeExpression.Value;
                            }

                            return string.Format("<{0}><FieldRef Name='{1}' LookupId=\'{4}\'/><Value Type='{2}'>{3}</Value></{0}>",
                            nodeExpression.Op.ToString(),
                            listFieldInfo.Name,
                            listFieldInfo.Type.ToString(),
                            value,
                            isID.ToString().ToUpper());
                        }
                        else if (listFieldInfo.Type == FieldType.Lookup && listFieldInfo.FieldTypeTaxonomy)
                        {
                            return string.Format("<{0}><FieldRef Name='{1}' /><Value Type='Text'>{2}</Value></{0}>",
                            nodeExpression.Op.ToString(),
                            listFieldInfo.Name,
                            value);
                        }
                        else if (listFieldInfo.Type == FieldType.User)
                        {
                            var spefObj = nodeExpression.Value as SPEFUser;
                            if (spefObj != null)
                                value = spefObj.ID;
                            else
                            {
                                if (nodeExpression.Value is int)
                                    value = nodeExpression.Value;
                            }

                            return string.Format("<{0}><FieldRef Name='{1}' LookupId='TRUE'/><Value Type='Integer'>{2}</Value></{0}>",
                                nodeExpression.Op.ToString(),
                                listFieldInfo.Name,
                                value);
                        }
                        else if (listFieldInfo.Type == FieldType.Boolean)
                        {
                            value = nodeExpression.Value ? "1" : "0";
                        }
                        else if (listFieldInfo.Type == FieldType.DateTime)
                        {
                            value = nodeExpression.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
                        }

                        return string.Format("<{0}><FieldRef Name='{1}'/><Value Type='{2}'>{3}</Value></{0}>",
                            nodeExpression.Op.ToString(),
                            listFieldInfo.Name,
                            listFieldInfo.Type.ToString(),
                            value);
                    }
                    else
                    {
                        return string.Format("<{0}><FieldRef Name='{1}'/></{0}>",
                            nodeExpression.Op.ToString(),
                            listFieldInfo.Name,
                            listFieldInfo.Type.ToString());
                    }
                }

                return string.Empty;
            }
            else
            {
                dynamic nodeOperation = node as SPEFOperation<T>;
                return string.Format("<{0}>{1}{2}</{0}>", nodeOperation.Operator.ToString(), getFilterQuery(nodeOperation.Operation1), getFilterQuery(nodeOperation.Operation2));
            }
        }

        private class NodeEval<T>
        {
            public NodeEval(SPEFOperation<T> node)
            {
                Operation = node;
            }
            public SPEFOperation<T> Operation { get; set; }
            public string Operation1 { get; set; }
            public string Operation2 { get; set; }
            public string GetValue()
            {
                return string.Format("<{0}>{1}{2}</{0}>", Operation.Operator.ToString(), Operation1, Operation2);
            }
        }
        private string getFilterQueryNonRec<T>(SPEFQueryNode<T> node)
        {
            var operationEvalList = new List<NodeEval<T>>();
            var t = node.GetType();
            if (t.IsGenericType && t.GetGenericTypeDefinition() == (typeof(SPEFExpression<,>)))
            {
                return getFilterQueryNodeNonRec(node);
            }
            else
            {
                dynamic nodeOperation = node as SPEFOperation<T>;
                operationEvalList.Add(new NodeEval<T>(nodeOperation));

                while (operationEvalList.Count > 0 && (string.IsNullOrWhiteSpace(operationEvalList[0].Operation1) || string.IsNullOrWhiteSpace(operationEvalList[0].Operation2)))
                {
                    var currentNode = operationEvalList.Last();

                    if (!string.IsNullOrWhiteSpace(currentNode.Operation1) && !string.IsNullOrWhiteSpace(currentNode.Operation2))
                    {
                        operationEvalList.Remove(currentNode);
                        var newLastNode = operationEvalList.Last();
                        if (string.IsNullOrWhiteSpace(newLastNode.Operation1))
                            newLastNode.Operation1 = currentNode.GetValue();
                        else if (string.IsNullOrWhiteSpace(newLastNode.Operation2))
                        {
                            newLastNode.Operation2 = currentNode.GetValue();
                        }
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(currentNode.Operation1))
                    {
                        var childNode = currentNode.Operation.Operation1;
                        var childType = childNode.GetType();
                        if (childType.IsGenericType && childType.GetGenericTypeDefinition() == (typeof(SPEFExpression<,>)))
                        {
                            var childOperationText = getFilterQueryNodeNonRec(childNode);
                            currentNode.Operation1 = childOperationText;
                        }
                        else
                        {
                            dynamic childNodeOperation = childNode as SPEFOperation<T>;
                            operationEvalList.Add(new NodeEval<T>(childNodeOperation));
                            continue;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(currentNode.Operation2))
                    {
                        var childNode = currentNode.Operation.Operation2;
                        var childType = childNode.GetType();
                        if (childType.IsGenericType && childType.GetGenericTypeDefinition() == (typeof(SPEFExpression<,>)))
                        {
                            var childOperationText = getFilterQueryNodeNonRec(childNode);
                            currentNode.Operation2 = childOperationText;
                        }
                        else
                        {
                            dynamic childNodeOperation = childNode as SPEFOperation<T>;
                            operationEvalList.Add(new NodeEval<T>(childNodeOperation));
                            continue;
                        }
                    }


                }

                return operationEvalList[0].GetValue();
            }
        }

        private string getFilterQueryNodeNonRec<T>(SPEFQueryNode<T> node)
        {
            var t = node.GetType();
            var exprTypes = t.GetGenericArguments();
            var t1 = exprTypes[1];
            var targetType = typeof(SPEFExpression<,>).MakeGenericType(typeof(T), t1);

            dynamic nodeExpression = Convert.ChangeType(node, targetType);

            var fieldInfo = getPropertyInfo(nodeExpression.Expression);

            var currListInfo = getSPEFStructureInfo<T>();

            var listFieldInfo = currListInfo.FieldsInfo.Where(f => f.PropertyInfo.Name == fieldInfo.Name).FirstOrDefault();
            if (listFieldInfo != null)
            {
                var op = (Op)nodeExpression.Op;
                if (!(op == Op.IsNotNull || op == Op.IsNull))
                {
                    var value = nodeExpression.Value;
                    if (listFieldInfo.Type == FieldType.Lookup && !listFieldInfo.FieldTypeTaxonomy)
                    {
                        var isID = true;
                        var spefObj = nodeExpression.Value as SPEFListItem;
                        if (spefObj != null)
                            value = spefObj.ID;
                        else
                        {
                            if (!(nodeExpression.Value is int))
                                isID = false;
                            value = nodeExpression.Value;
                        }

                        return string.Format("<{0}><FieldRef Name='{1}' LookupId=\'{4}\'/><Value Type='{2}'>{3}</Value></{0}>",
                        nodeExpression.Op.ToString(),
                        listFieldInfo.Name,
                        listFieldInfo.Type.ToString(),
                        value,
                        isID.ToString().ToUpper());
                    }
                    else if (listFieldInfo.Type == FieldType.Lookup && listFieldInfo.FieldTypeTaxonomy)
                    {
                        return string.Format("<{0}><FieldRef Name='{1}' /><Value Type='Text'>{2}</Value></{0}>",
                        nodeExpression.Op.ToString(),
                        listFieldInfo.Name,
                        value);
                    }
                    else if (listFieldInfo.Type == FieldType.User)
                    {
                        var spefObj = nodeExpression.Value as SPEFUser;
                        if (spefObj != null)
                            value = spefObj.ID;
                        else
                        {
                            if (nodeExpression.Value is int)
                                value = nodeExpression.Value;
                        }

                        return string.Format("<{0}><FieldRef Name='{1}' LookupId='TRUE'/><Value Type='Integer'>{2}</Value></{0}>",
                            nodeExpression.Op.ToString(),
                            listFieldInfo.Name,
                            value);
                    }
                    else if (listFieldInfo.Type == FieldType.Boolean)
                    {
                        value = nodeExpression.Value ? "1" : "0";
                    }
                    else if (listFieldInfo.Type == FieldType.DateTime)
                    {
                        value = nodeExpression.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
                    }

                    return string.Format("<{0}><FieldRef Name='{1}'/><Value Type='{2}'>{3}</Value></{0}>",
                        nodeExpression.Op.ToString(),
                        listFieldInfo.Name,
                        listFieldInfo.Type.ToString(),
                        value);
                }
                else
                {
                    return string.Format("<{0}><FieldRef Name='{1}'/></{0}>",
                        nodeExpression.Op.ToString(),
                        listFieldInfo.Name,
                        listFieldInfo.Type.ToString());
                }
            }

            return string.Empty;
        }

        private string getSortQuery<T, TP2>(SPEFSortNode<T, TP2> orderBy, out SPEFFieldInfo orderListFieldInfo)
        {
            var fieldInfo = getPropertyInfo(orderBy.Expression);

            var currListInfo = getSPEFStructureInfo<T>();

            var listFieldInfo = currListInfo.FieldsInfo.Where(f => f.PropertyInfo.Name == fieldInfo.Name).FirstOrDefault();
            orderListFieldInfo = listFieldInfo;
            if (listFieldInfo != null)
            {
                var filterQuery = getFilterQueryNonRec(orderBy.Query);
                if (!string.IsNullOrWhiteSpace(filterQuery))
                    filterQuery = string.Format("<Where>{0}</Where>", filterQuery);
                return string.Format("{0}<OrderBy><FieldRef Name='{1}' Ascending='{2}'/></OrderBy>", filterQuery, listFieldInfo.Name, orderBy.Ascending.ToString());
            }

            return string.Empty;
        }



        private PropertyInfo getPropertyInfo<TSource, TProperty>(Expression<Func<TSource, TProperty>> propertyLambda)
        {
            Type type = typeof(TSource);

            MemberExpression member = propertyLambda.Body as MemberExpression;
            if (member == null)
                throw new ArgumentException(string.Format(
                    "Expression '{0}' refers to a method, not a property.",
                    propertyLambda.ToString()));

            PropertyInfo propInfo = member.Member as PropertyInfo;
            if (propInfo == null)
                throw new ArgumentException(string.Format(
                    "Expression '{0}' refers to a field, not a property.",
                    propertyLambda.ToString()));

            if (type != propInfo.ReflectedType &&
                !type.IsSubclassOf(propInfo.ReflectedType))
                throw new ArgumentException(string.Format(
                    "Expresion '{0}' refers to a property that is not from type {1}.",
                    propertyLambda.ToString(),
                    type));

            return propInfo;
        }

        private List<T> loadInto<T>(ListItemCollection listItemCollection) where T : SPEFListItem
        {
            var retList = new List<T>();
            foreach (var listItem in listItemCollection)
            {
                var obj = loadInto<T>(listItem);
                retList.Add(obj);
            }
            return retList;
        }

        private T loadInto<T>(ListItem listItem) where T : SPEFListItem
        {
            var errors = new StringBuilder();
            var currListInfo = getSPEFStructureInfo<T>();

            var retObj = (T)Activator.CreateInstance(typeof(T));
            retObj.ID = listItem.Id;

            foreach (var fieldInfo in currListInfo.FieldsInfo)
            {
                try
                {
                    var fieldValue = getFieldValue(listItem, fieldInfo);
                    fieldInfo.PropertyInfo.SetValue(retObj, fieldValue);
                }
                catch (Exception ex)
                {
                    errors.AppendFormat("- {0} ({1}) ", fieldInfo.Name, ex.Message);
                }

                var errorsString = errors.ToString();
                if (!string.IsNullOrWhiteSpace(errorsString))
                {
                    throw new Exception(string.Format("Error loading item ID {0} from list {1} into SPEFListItem object - {2}", listItem.Id, currListInfo.Title, errorsString));
                }
            }

            if (listItem.AttachmentFiles != null && listItem.AttachmentFiles.AreItemsAvailable)
            {
                foreach (var spAttachment in listItem.AttachmentFiles)
                {
                    var attachment = new SPEFAttachment();
                    attachment.FileName = spAttachment.FileName;
                    attachment.ServerRelativeUrl = spAttachment.ServerRelativeUrl;
                    retObj.Attachments.Add(attachment);
                }
            }

            //retObj.SetInit();
            return retObj;
        }


        private ContentTypeId getContentTypeID(string contentTypeName)
        {
            using (var mainContext = GetSharePointContext())
            {
                var webContentTypeCollection = mainContext.Web.ContentTypes;
                mainContext.Load(webContentTypeCollection, ct => ct.Include(l => l.Name).Where(l => l.Name == contentTypeName));
                mainContext.ExecuteQuery();

                ContentType contentType = null;
                if (webContentTypeCollection.Count > 0)
                {
                    contentType = webContentTypeCollection[0];
                    return contentType.Id;
                }
            }
            return null;
        }

        #endregion

        #region Delete Objects

        public bool Delete<T>(T entity, string label = null) where T : SPEFListItem
        {
            return Delete<T>(entity.ID, label);
        }

        public bool Delete<T>(int id, string label = null) where T : SPEFListItem
        {
            if (id <= 0)
                return false;

            var currListInfo = getSPEFStructureInfo<T>();

            using (var context = GetSharePointContext(label ?? currListInfo.ContextUrl))
            {
                var web = context.Web;
                var listCollection = web.Lists;
                var listTitle = currListInfo.Title;
                context.Load(listCollection, lists => lists.Include(l => l.Title).Where(l => l.Title == listTitle));
                context.ExecuteQuery();

                List list = null;
                if (listCollection.Count == 0)
                    return false;
                list = listCollection[0];

                var currListItem = list.GetItemById(id);
                context.Load(currListItem);
                try
                {
                    context.ExecuteQuery();
                    currListItem.DeleteObject();
                    context.ExecuteQuery();
                }
                catch (Exception exNotFound)
                {
                    return false;
                }
                return true;
            }
        }

        public bool Delete<T>(SPEFQueryNode<T> node, string label = null) where T : SPEFListItem
        {
            var currListInfo = getSPEFStructureInfo<T>();
            var spQuery = string.Empty;
            if (node != null)
            {
                SPEFFieldInfo sortListFieldInfo = null;
                spQuery = getQuery(node, out sortListFieldInfo);
            }
            var xmlQuery = string.Format("<View><Query>{0}</Query></View>", spQuery);

            using (var context = GetSharePointContext(label ?? currListInfo.ContextUrl))
            {
                try
                {
                    var web = context.Web;
                    var listCollection = web.Lists;
                    var listTitle = currListInfo.Title;
                    context.Load(listCollection, lists => lists.Include(l => l.Title).Where(l => l.Title == listTitle));
                    context.ExecuteQuery();

                    if (listCollection.Count == 0)
                        return false;

                    var list = listCollection[0];

                    var camlQuery = new CamlQuery() { ViewXml = xmlQuery };

                    var listItemCollection = list.GetItems(camlQuery);
                    context.Load(listItemCollection);

                    context.ExecuteQuery();
                    foreach (var item in listItemCollection)
                        item.DeleteObject();

                    context.ExecuteQuery();
                    return true;
                }
                catch (Exception exNotFound)
                {
                    return false;
                }
            }
        }

        #endregion

        #region Save Objects

        public int Save<T>(T entity, string label = null, SPEFUser editor = null, SPEFUser author = null) where T : SPEFListItem
        {
            if (editor != null)
                entity.SetEditor(editor);

            if (author != null)
                entity.SetAuthor(author);
            else
            {
                if (entity.ID <= 0 && !entity.AuthorUpdated && editor != null)
                {
                    entity.SetAuthor(editor);
                }
            }

            var typeName = entity.GetType().Name; //typeof(T).Name;

            var currKey = string.Format(STRUCTS_KEY, typeName);
            var currListInfo = structsMapping[currKey];

            using (var context = GetSharePointContext(label ?? currListInfo.ContextUrl))
            {
                var web = context.Web;
                var listCollection = web.Lists;
                var listTitle = currListInfo.Title;
                context.Load(listCollection, lists => lists.Include(l => l.Title, l => l.EnableAttachments).Where(l => l.Title == listTitle));
                context.ExecuteQuery();

                List list = null;
                if (listCollection.Count == 0)
                    return -1;
                list = listCollection[0];

                ListItem currListItem = null;
                if (entity.ID > 0)
                {
                    currListItem = list.GetItemById(entity.ID);
                    context.Load(currListItem);
                    context.ExecuteQuery();
                }
                if (currListItem == null)
                {
                    if (currListInfo.TemplateType == ListTemplateType.DocumentLibrary)
                    {
                        var docEntity = entity as SPEFDocListItem;
                        if (docEntity != null)
                        {
                            //create new document
                            var fileCreationInformation = new FileCreationInformation();
                            fileCreationInformation.Content = docEntity.FileContent;
                            fileCreationInformation.Overwrite = docEntity.FileOverwrite;
                            if (string.IsNullOrWhiteSpace(docEntity.FileUrl))
                                docEntity.FileUrl = docEntity.FileName;
                            fileCreationInformation.Url = docEntity.FileUrl;

                            File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
                            context.Load(uploadFile, f => f.ListItemAllFields);
                            currListItem = uploadFile.ListItemAllFields;
                        }
                    }
                    else
                    {
                        //create new item
                        var listItemCreationInformation = new ListItemCreationInformation();
                        currListItem = list.AddItem(listItemCreationInformation);
                    }
                    context.ExecuteQuery();
                }
                if (list.EnableAttachments)
                {
                    context.Load(currListItem.AttachmentFiles);
                    context.ExecuteQuery();

                    var modified = false;

                    for (int i = currListItem.AttachmentFiles.Count - 1; i >= 0; i--)
                    {
                        var prevAttachment = currListItem.AttachmentFiles[i];
                        //se l'allegato non è più presente, lo rimuovo
                        if (!entity.Attachments.Exists(a => a.FileName == prevAttachment.FileName))
                        {
                            currListItem.AttachmentFiles[i].DeleteObject();
                            modified = true;
                        }
                    }
                    foreach (var attachment in entity.Attachments)
                    {
                        //se l'allegato era già presente, vado avanti
                        if (currListItem.AttachmentFiles.Where(a => a.FileName == attachment.FileName).FirstOrDefault() != null)
                        {
                            continue;
                        }

                        var attInfo = new AttachmentCreationInformation();
                        attInfo.FileName = attachment.FileName;
                        attInfo.ContentStream = new System.IO.MemoryStream(attachment.FileContent);

                        Attachment att = currListItem.AttachmentFiles.Add(attInfo);
                        modified = true;
                    }
                    if (modified)
                    {
                        context.ExecuteQuery();
                        context.Load(currListItem);
                        context.ExecuteQuery();
                    }
                }

                #region fields
                foreach (var fieldInfo in currListInfo.FieldsInfo)
                {
                    if (fieldInfo.Readonly)
                        continue;
                    if (fieldInfo.Name == StandardListItemFields.Author && !entity.AuthorUpdated)
                        continue;
                    if (fieldInfo.Name == StandardListItemFields.Editor && !entity.EditorUpdated)
                        continue;
                    try
                    {
                        var fieldValue = getPropertyValue(entity, fieldInfo);
                        if (fieldValue == null && fieldInfo.Ignore)
                            continue;
                        setFieldValue(currListItem, fieldInfo, fieldValue);
                    }
                    catch (Exception ex)
                    {
                        return -1;
                    }
                }
                #endregion

                #region attachments
                if (list.EnableAttachments)
                {

                }
                #endregion

                currListItem.Update();
                context.ExecuteQuery();
                return currListItem.Id;
            }
        }

        private object getPropertyValue(SPEFListItem listItemEntity, SPEFFieldInfo fieldInfo)
        {
            return fieldInfo.PropertyInfo.GetValue(listItemEntity);
        }

        private TimeZoneInfo serverTimeZoneInfo;
        private TimeZoneInfo ServerTimeZoneInfo
        {
            get
            {
                if (serverTimeZoneInfo == null)
                {
                    using (var context = GetSharePointContext())
                    {

                        //1.Retrieve SharePoint TimeZone
                        var spTimeZone = context.Web.RegionalSettings.TimeZone;
                        context.Load(spTimeZone);
                        context.ExecuteQuery();

                        //2.Resolve System.TimeZoneInfo from Microsoft.SharePoint.Client.TimeZone 
                        var fixedTimeZoneName = spTimeZone.Description.Replace("and", "&");
                        serverTimeZoneInfo = TimeZoneInfo.GetSystemTimeZones().FirstOrDefault(tz => tz.DisplayName == fixedTimeZoneName);
                    }
                }
                return serverTimeZoneInfo;
            }

        }

        private DateTime ConvertToServerDate(DateTime date)
        {
            try
            {
                return TimeZoneInfo.ConvertTime(date, ServerTimeZoneInfo);
            }
            catch (Exception ex)
            {
                return date;
            }
        }

        private void setFieldValue(ListItem listItem, SPEFFieldInfo fieldInfo, object value)
        {
            if (value == null)
                listItem.SetValue(fieldInfo.Name, null);

            switch (fieldInfo.Type)
            {
                case FieldType.MultiChoice:
                    var choices = value as string[];
                    listItem.SetMultiChoiceValue(fieldInfo.Name, choices);
                    return;
                case FieldType.Lookup:
                    if (fieldInfo.FieldTypeTaxonomy)
                    {
                        if (fieldInfo.Multiple)
                        {
                            var termsSPEF = value as List<SPEFTaxonomyItem>;
                            if (termsSPEF != null)
                            {
                                listItem.SetSPEFMultiTaxonomyValues(fieldInfo.Name, termsSPEF);
                            }
                            else
                            {
                                var termsKeyValue = value as List<KeyValuePair<string, string>>;
                                if (termsKeyValue != null)
                                    listItem.SetMultiTaxonomyValues(fieldInfo.Name, termsKeyValue);
                            }
                        }
                        else
                        {
                            var termSPEF = value as SPEFTaxonomyItem;
                            if (termSPEF != null)
                            {
                                listItem.SetSPEFTaxonomyValue(fieldInfo.Name, termSPEF);
                            }
                            else
                            {
                                var termKeyValue = (KeyValuePair<string, string>)value;
                                listItem.SetTaxonomyValue(fieldInfo.Name, termKeyValue);
                            }
                        }
                    }
                    else
                    {
                        if (fieldInfo.Multiple)
                        {
                            IEnumerable list = value as IEnumerable;
                            if (list == null)
                                listItem.SetValue(fieldInfo.Name, null);

                            var ids = new List<int>();
                            foreach (var o in list)
                            {
                                var refObject = o as SPEFItem;
                                if (refObject != null)
                                    ids.Add(refObject.ID);
                            }
                            listItem.SetMultiLookupIdValues(fieldInfo.Name, ids);
                        }
                        else
                        {
                            var refObject = value as SPEFItem;
                            if (refObject != null && refObject.ID > 0)
                                listItem.SetLookupIdValue(fieldInfo.Name, refObject.ID);
                        }
                    }
                    return;
                case FieldType.URL:
                    listItem.SetUrlValue(fieldInfo.Name, (value ?? "").ToString());
                    return;
                case FieldType.DateTime:
                    var date = DateTime.MinValue;
                    if (DateTime.TryParse(value.ToString(), out date) && date != DateTime.MinValue)
                    {
                        //var utcValue = date.ToUniversalTime();
                        value = ConvertToServerDate(date); //date; //ConvertToServerDate(date);
                        listItem.SetValue(fieldInfo.Name, value);
                    }
                    return;
                case FieldType.User:
                    if (!fieldInfo.Multiple)
                    {
                        var user = value as SPEFUser;
                        if (user != null && user.ID > 0)
                            listItem.SetUserValue(fieldInfo.Name, user.ID);
                    }
                    else
                    {
                        var users = value as List<SPEFUser>;
                        if (users != null)
                            listItem.SetMultiUserValue(fieldInfo.Name, users.Select(u => u.AccountName).ToArray());
                    }
                    return;
                default:
                    listItem.SetValue(fieldInfo.Name, value);
                    return;
            }
        }

        #endregion

        #region Taxonomy

        public List<SPEFTaxonomyItem> GetTaxonomyItems(string termSetName, int lcid) //, string termStoreName, string termGroupName
        {
            var termStoreID = Guid.Empty;
            var termGroupID = Guid.Empty;
            var termSetID = Guid.Empty;



            using (var context = GetSharePointContext())
            {
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

                context.Load(taxonomySession,
                    ts => ts.TermStores.Include(
                        store => store.Name,
                        store => store.Id,
                        store => store.Groups.Include(
                            group => group.Name,
                            group => group.Id,
                            group => group.TermSets.Include(
                                set => set.Name,
                                set => set.Id
                            )
                        )
                    )
                );

                context.ExecuteQuery();

                if (taxonomySession == null)
                {
                    return null;
                }

                TermSet termSet = null;
                foreach (var store in taxonomySession.TermStores)
                {
                    foreach (var group in store.Groups)
                    {
                        foreach (var set in group.TermSets)
                        {
                            if (set.Name == termSetName)
                            {
                                termSet = set;
                            }
                        }
                    }
                }
                if (termSet == null)
                {
                    return new List<SPEFTaxonomyItem>();
                }

                context.Load(termSet, ts => ts.Terms.Include(t => t.TermsCount, t => t.Id, t => t.Name));
                context.ExecuteQuery();

                if (termSet == null)
                {
                    return new List<SPEFTaxonomyItem>();
                }

                var retList = new List<SPEFTaxonomyItem>();
                foreach (var term in termSet.Terms)
                {
                    retList.Add(new SPEFTaxonomyItem() { ID = term.Id, Value = term.Name });
                }

                return retList;
            }
        }

        public SPEFTaxonomyItem CheckTaxonomyItem(string termSetName, string value)
        {
            using (var context = GetSharePointContext())
            {
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

                context.Load(taxonomySession,
                    ts => ts.TermStores.Include(
                        store => store.Name,
                        store => store.Id,
                        store => store.Groups.Include(
                            group => group.Name,
                            group => group.Id,
                            group => group.TermSets.Include(
                                set => set.Name,
                                set => set.Id
                            )
                        )
                    )
                );

                context.ExecuteQuery();

                if (taxonomySession == null)
                {
                    return null;
                }

                TermSet termSet = null;
                foreach (var store in taxonomySession.TermStores)
                {
                    foreach (var group in store.Groups)
                    {
                        foreach (var set in group.TermSets)
                        {
                            if (set.Name == termSetName)
                            {
                                termSet = set;
                            }
                        }
                    }
                }
                if (termSet == null)
                {
                    return null;
                }

                context.Load(termSet, ts => ts.Terms.Include(t => t.TermsCount, t => t.Id, t => t.Name));
                context.ExecuteQuery();

                if (termSet == null)
                {
                    return null;
                }

                var valueToCheck = value.Trim().ToLower();
                var term = termSet.Terms.Where(t => t.Name.ToLower() == valueToCheck).FirstOrDefault();
                if (term != null)
                    return new SPEFTaxonomyItem() { ID = term.Id, Value = term.Name };

                var termID = Guid.NewGuid();
                var retTerm = termSet.CreateTerm(value, 1033, termID);
                context.ExecuteQuery();

                return new SPEFTaxonomyItem() { ID = termID, Value = value };
            }
        }

        #endregion

        #region Users

        private User LoadUser(int id)
        {
            try
            {
                using (var context = GetSharePointContext())
                {
                    var user = context.Web.GetUserById(id);
                    return user;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public bool IsUser(int id)
        {
            using (var context = GetSharePointContext())
            {
                try
                {
                    var userInfoList = context.Site.RootWeb.SiteUserInfoList;
                    var userInfo = userInfoList.GetItemById(id);
                    context.Load(userInfo, i => i.ContentType);
                    context.ExecuteQuery();
                    return userInfo.ContentType.Name == "Person";
                }
                catch (Exception ex)
                {
                    return true;
                }
            }
        }

        public bool IsUserInGroup(int userID, string groupName)
        {
            using (var context = GetSharePointContext())
            {
                var collGroup = context.Web.SiteGroups;
                var oGroup = collGroup.GetByName(groupName);
                var collUser = oGroup.Users;

                context.Load(collUser);
                context.ExecuteQuery();

                foreach (var oUser in collUser)
                {
                    if (oUser.Id == userID)
                        return true;
                }
                return false;
            }
        }

        public List<SPEFUser> LoadAllUserWithInfo()
        {
            var allGroups = LoadAllGroups();

            var retUsers = new List<SPEFUser>();
            try
            {
                using (var context = GetSharePointContext())
                {
                    var userProfilesResult = new List<PersonProperties>();
                    var web = context.Web;
                    var peopleManager = new PeopleManager(context);


                    var siteUsers = from user in web.SiteUsers
                                    where user.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User
                                    select user;
                    var usersResult = context.LoadQuery(siteUsers);
                    context.ExecuteQuery();

                    var userResultList = usersResult.ToList();

                    var count = 0;
                    foreach (var user in userResultList)
                    {
                        count++;

                        context.Load(user.Groups);

                        var userProfile = peopleManager.GetPropertiesFor(user.LoginName);
                        context.Load(userProfile,
                            p => p.AccountName,
                            p => p.UserProfileProperties,
                            p => p.DisplayName,
                            p => p.Email,
                            p => p.PersonalUrl,
                            p => p.PictureUrl);
                        userProfilesResult.Add(userProfile);

                        if (count == 10)
                        {
                            context.ExecuteQuery();
                            count = 0;
                        }
                    }
                    if (count > 0)
                        context.ExecuteQuery();

                    for (int i = 0; i < userResultList.Count; i++)
                    {
                        var user = userResultList[i];
                        var personProperties = userProfilesResult[i];

                        if (!personProperties.ServerObjectIsNull.HasValue || personProperties.ServerObjectIsNull.Value)
                            continue;

                        var existingsUser = retUsers.Where(u => u.AccountName == personProperties.AccountName).FirstOrDefault();
                        if (existingsUser != null)
                        {
                            existingsUser.IDs.Add(user.Id);
                            continue;
                        }

                        var retInfo = new SPEFUser(user.Id);
                        retInfo.Email = user.Email;
                        retInfo.IsSiteAdmin = user.IsSiteAdmin;
                        retInfo.AccountName = personProperties.AccountName;
                        retInfo.UserUrl = personProperties.PersonalUrl;
                        retInfo.PictureUrl = personProperties.PictureUrl;
                        retInfo.DisplayName = personProperties.DisplayName;
                        //retInfo.HasLoginType = user.LoginName.Contains('#') && user.LoginName.Contains('|');

                        // User Profile Properties
                        if (personProperties.UserProfileProperties != null)
                        {
                            var properties = personProperties.UserProfileProperties;
                            retInfo.Properties = properties;
                        }
                        foreach (var group in user.Groups)
                        {
                            var currGroup = allGroups.Where(g => g.ID == group.Id).FirstOrDefault();
                            if (currGroup != null)
                                retInfo.Groups.Add(currGroup);
                        }
                        retUsers.Add(retInfo);
                    }
                }
                return retUsers;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        protected int GetUserIdFromLoginName(string loginName)
        {
            // set up domain context
            using (var context = GetSharePointContext())
            {

                context.Load(context.Site, l => l.RootWeb.SiteUsers, l => l.RootWeb.Id, l => l.Id);
                context.ExecuteQuery();

                var user = context.Site.RootWeb.SiteUsers.GetByLoginName(loginName);
                context.Load(user, l => l.Groups, l => l.Id, l => l.LoginName);
                context.ExecuteQuery();

                if (user != null)
                    return user.Id;

                return -1;
            }
        }

        public SPEFUser LoadUserWithInfo(string loginName)
        {
            var userID = GetUserIdFromLoginName(loginName);
            if (userID >= 0)
                return LoadUserWithInfo(userID);
            return null;
        }

        public SPEFUser LoadUserWithInfo(int id)
        {
            var isUser = true;
            try
            {
                isUser = IsUser(id);

                using (var context = GetSharePointContext())
                {
                    if (isUser)
                    {
                        var user = context.Web.GetUserById(id);

                        if (user != null)
                        {
                            var peopleManager = new PeopleManager(context);

                            context.Load(user,
                                    u => u.Id,
                                    u => u.UserId,
                                    u => u.LoginName,
                                    u => u.Email,
                                    u => u.IsSiteAdmin,
                                    u => u.Title,
                                    u => u.Groups);
                            context.ExecuteQuery();

                            PersonProperties personProperties = peopleManager.GetPropertiesFor(user.LoginName);

                            context.Load(personProperties,
                                p => p.AccountName,
                                p => p.UserProfileProperties,
                                p => p.DisplayName,
                                p => p.Email,
                                p => p.PersonalUrl,
                                p => p.PictureUrl);
                            context.ExecuteQuery();

                            var retInfo = new SPEFUser(user.Id);
                            retInfo.IDs.Add(user.Id);
                            retInfo.Email = user.Email;
                            retInfo.IsSiteAdmin = user.IsSiteAdmin;
                            retInfo.AccountName = personProperties.AccountName;
                            retInfo.UserUrl = personProperties.PersonalUrl;
                            retInfo.PictureUrl = personProperties.PictureUrl;
                            retInfo.DisplayName = personProperties.DisplayName;
                            // retInfo.HasLoginType = user.LoginName.Contains('#') && user.LoginName.Contains('|');

                            // User Profile Properties
                            if (personProperties.UserProfileProperties != null)
                            {
                                var properties = personProperties.UserProfileProperties;
                                retInfo.Properties = properties;
                            }

                            return retInfo;
                        }

                    }
                    else
                    {
                        var retGroup = new SPEFUser(id);
                        retGroup.IsGroup = true;

                        var siteGroups = context.Web.SiteGroups;
                        var membersGroup = siteGroups.GetById(id);
                        context.Load(membersGroup);
                        context.Load(membersGroup.Users, mUsers => mUsers.Include(
                                u => u.Id,
                                u => u.UserId,
                                u => u.LoginName,
                                u => u.Email,
                                u => u.IsSiteAdmin,
                                u => u.Title,
                                u => u.Groups));
                        context.ExecuteQuery();

                        retGroup.DisplayName = membersGroup.Title;
                        retGroup.AccountName = membersGroup.LoginName;

                        foreach (var member in membersGroup.Users)
                        {
                            retGroup.Members.Add(
                                new SPEFUser(member.Id)
                                {
                                    DisplayName = member.Title,
                                    Email = member.Email,
                                    AccountName = member.LoginName,
                                });
                        }
                        return retGroup;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public List<SPEFUser> LoadAllGroups()
        {
            var retGroups = new List<SPEFUser>();
            try
            {
                using (var context = GetSharePointContext())
                {
                    var web = context.Web;

                    var siteGroups = from g in web.SiteGroups
                                     select g;
                    var groupsResult = context.LoadQuery(siteGroups);
                    context.ExecuteQuery();

                    var groupsResultList = groupsResult.ToList();

                    for (int i = 0; i < groupsResultList.Count; i++)
                    {
                        var group = groupsResultList[i];

                        var retInfo = new SPEFUser(group.Id);
                        retInfo.Email = string.Empty;
                        retInfo.AccountName = group.LoginName;
                        retInfo.DisplayName = group.Title;
                        retInfo.IsGroup = true;

                        retGroups.Add(retInfo);
                    }
                }
                return retGroups;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        #endregion

        #region Populate lists
        public List<string> Populate()
        {
            var retMessages = new List<string>();
            try
            {
                foreach (var listInfo in structsMapping.Values.Where(s => s.IsList))
                {
                    if (populateList(listInfo))
                        retMessages.Add(string.Format("List OK: {0}", listInfo.Title));
                }
            }
            catch (Exception ex)
            {
                retMessages.Add(string.Format("ERROR: {0}", ex.Message));
                return retMessages;
            }
            return retMessages;
        }

        private bool populateList(SPEFStructInfo listInfo, string contextUrl = null)
        {
            using (var context = GetSharePointContext(contextUrl ?? listInfo.ContextUrl))
            {
                var web = context.Web;

                var listCollection = web.Lists;
                var listTitle = listInfo.Title;
                context.Load(listCollection, lists => lists.Include(l => l.Title).Where(l => l.Title == listTitle));
                context.ExecuteQuery();

                List list = null;
                if (listCollection.Count > 0)
                    list = listCollection[0];
                else
                {
                    return false;
                }

                var items = list.GetItems(new CamlQuery());
                context.Load(items);
                context.ExecuteQuery();
                if (items.Count > 0)
                    return true;

                for (int i = 0; i < 10; i++)
                {
                    var retObj = Convert.ChangeType(Activator.CreateInstance(listInfo.StructType), listInfo.StructType);

                    foreach (var fieldInfo in listInfo.FieldsInfo)
                    {
                        if (fieldInfo.Hidden || fieldInfo.Readonly)
                            continue;
                        try
                        {
                            var fieldValue = getRandomFieldValue(fieldInfo);
                            fieldInfo.PropertyInfo.SetValue(retObj, fieldValue);
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }

                    typeof(SPEFContext)
                        .GetMethod("Save")
                        .MakeGenericMethod(listInfo.StructType)
                        .Invoke(this, new[] { retObj, null, null, null });
                }
            }
            return true;
        }

        Random randomFieldExtractor = new Random();
        const string LoremIpsum = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.";
        private object getRandomFieldValue(SPEFFieldInfo fieldInfo)
        {
            switch (fieldInfo.Type)
            {
                case FieldType.Note:
                case FieldType.Text:
                    var sb = new StringBuilder();
                    for (int i = 0; i < 5; i++)
                        sb.Append(randomFieldExtractor.Next(10));

                    var thisString = string.Format($"{sb} {LoremIpsum}");
                    var length = randomFieldExtractor.Next(Math.Min(fieldInfo.MaxLength, thisString.Length));
                    return thisString.Substring(0, length);

                case FieldType.Boolean:
                    var retBool = randomFieldExtractor.NextDouble() < 0.5;
                    return retBool;
                case FieldType.Number:
                    var retNumber = randomFieldExtractor.Next(0, 1000);
                    return retNumber;
                case FieldType.Currency:
                    var retCurrency = randomFieldExtractor.Next(0, 1000);
                    return retCurrency;
                case FieldType.DateTime:
                    var days = randomFieldExtractor.Next(-1000, 1000);
                    var retDate = DateTime.Now.Date.AddDays(days);
                    return retDate;
                case FieldType.Choice:
                    var randomChoice = randomFieldExtractor.Next(0, fieldInfo.Choices.Count);
                    var selChoice = fieldInfo.Choices[randomChoice];
                    return selChoice;
                case FieldType.MultiChoice:
                    {
                        var randomChoice1 = randomFieldExtractor.Next(0, fieldInfo.Choices.Count);
                        var selChoice1 = fieldInfo.Choices[randomChoice1];
                        return new string[] { selChoice1 };
                    }
                case FieldType.Lookup:
                    {
                        if (fieldInfo.FieldTypeTaxonomy)
                        {
                            return null;
                        }
                        else
                        {
                            object loadedObjects = null;
                            if (!fieldInfo.Multiple)
                            {
                                var method1 = typeof(SPEFContext).GetMethods().Where(c => c.Name == "Load").ToList()[1];
                                var method = method1.MakeGenericMethod(fieldInfo.PropertyInfo.PropertyType);

                                loadedObjects = method.Invoke(this, new object[] { null, null });
                            }
                            else
                            {
                                var method1 = typeof(SPEFContext).GetMethods().Where(c => c.Name == "Load").ToList()[1];
                                var method = method1.MakeGenericMethod(fieldInfo.PropertyInfo.PropertyType.GenericTypeArguments[0]);

                                loadedObjects = method.Invoke(this, new object[] { null, null });
                            }
                            Type generic = typeof(List<>);
                            Type[] typeArgs = { fieldInfo.PropertyInfo.PropertyType };
                            Type constructed = generic.MakeGenericType(typeArgs);

                            var spefObjects = (IList)Convert.ChangeType(loadedObjects, constructed);
                            if (spefObjects.Count == 0)
                                return null;

                            var refObjItem = spefObjects[0] as SPEFListItem;
                            if (!fieldInfo.Multiple)
                            {
                                return refObjItem;
                            }
                            else
                            {
                                var refObjType = fieldInfo.PropertyInfo.PropertyType;
                                var refObj = Activator.CreateInstance(refObjType);
                                refObjType.GetMethod("Add").Invoke(refObj, new[] { refObjItem });
                                return refObj;
                            }

                        }
                    }
                case FieldType.URL:
                    {
                        return "http://www.google.it";
                    }
                case FieldType.User:
                    return null;
                default:
                    return null;
            }
        }

        #endregion
    }
}
