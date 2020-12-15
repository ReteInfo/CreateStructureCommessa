using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.SharePoint.Client;
using SPClient = Microsoft.SharePoint.Client;
using Tax =  Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.DocumentManagement;
using Newtonsoft.Json;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Text.RegularExpressions;
using System.ComponentModel;
using Microsoft.Extensions.Logging;

namespace Company.Function
{
    static class Documents
    {
        private static string _MMS = "08849a00032a4803a870e0addbaf5a9c";
        public static string siteUrl = "https://reteinformatica.sharepoint.com/sites/";
        
        private static string _TermGroup = "22e57a67-4f40-4463-8ec6-21f17af008e1";
        private static string _GroupColumn = "Unifor";
        private static string _nameContentType = "UniforDoc";
        private static string _ListConfig = "Configurations";
        private static string _ListLibraryName = "Documents";
        private static string _internalNameList = "Shared Documents";
        
        
        

        private static Dictionary<string, string> columnContentType = new Dictionary<string, string>()
          {
              { "CompaniesMT","21604c6a-be83-4319-b86e-f3e951a3b888,Canale" },
              { "PMsMT", "81707b84-968f-4844-b525-a761a2bc3c38,PM" },
              { "CustomersMT","888224c8-2372-4219-8c57-1810a21a1d94,Project Name"},
              { "YearsMT","35f5f65e-fd60-4baf-ab25-e12e67993c09,Anno"},
              { "ISOMT","20e7347a-4463-4eca-a09d-9f3d184b75c0,ISO"},
              { "NumberOfferMT","28ee3c2e-666f-41d1-902c-8869eba5b0ed,Numero Offerta"},
              { "DocTypesMT","c2b00a5a-d819-43b1-a240-821043f1bc62,Tipo Documento"},
              { "SubDocTypesMT","292f3171-9035-491f-93f1-ed4838871494,Sub Tipo Documento"},
              { "CitiesMT","49891502-7d91-4e5d-9e51-e55b0646583f,Citta"},
          };

   
    public static string GetDefaultFolderValues(SPClient.List sourceList,ILogger log)
       {
           try
           {
               var sourceContext = (ClientContext)sourceList.Context;

                SPClient.Folder formsFolder =
                sourceContext.Web.GetFolderByServerRelativeUrl(sourceList.RootFolder.ServerRelativeUrl + "/forms");

                sourceContext.Load(formsFolder, f => f.Files);
                sourceContext.ExecuteQuery();

                SPClient.File clientLocationBasedDefaultsFile =
                    formsFolder.Files.FirstOrDefault(
                        f => f.Name.ToLowerInvariant() == "client_LocationBasedDefaults.html".ToLowerInvariant());

                if (clientLocationBasedDefaultsFile != null)
                {
                    return ReadFileContent(clientLocationBasedDefaultsFile);
                }
                return null;
           }
           catch (System.Exception e)
           {
               log.LogError("Errore recupero deafult folder, source list {0}: {1}",sourceList,e.Message);
               throw e;
           }
      }

        public static string ReadFileContent(SPClient.File file)
        {
            ClientResult<Stream> stream = file.OpenBinaryStream();
            file.Context.ExecuteQuery();

            using (StreamReader reader = new StreamReader(stream.Value, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }


        public static string getTaxonomyTermGroup(ClientContext clientContext,string termName, ILogger log)
        {
            try
            {
                // Get the TaxonomySession
                Tax.TaxonomySession taxonomySession = Tax.TaxonomySession.GetTaxonomySession(clientContext);
                //MMS
                Guid guidStore = new Guid(_MMS);
                Tax.TermStore termStore = taxonomySession.TermStores.GetById(guidStore);
                //Guid Term Group - cellini
                System.Guid guidPeople = new System.Guid(_TermGroup);
                Tax.TermGroup termGroup = termStore.GetGroup(guidPeople);
                //get term set - Sezioni,Building,...
                Tax.TermSetCollection termSetColl = termGroup.TermSets;
                clientContext.Load(termSetColl);
                // Execute the query to the server
                clientContext.ExecuteQuery();

                var GuidTerm = "";
                // Loop through all the termsets
                foreach (Tax.TermSet termSet in termSetColl)
                {
                    //Console.WriteLine(termSet.Name);
                    Tax.LabelMatchInformation termQuery = new Tax.LabelMatchInformation(clientContext)
                    {
                        TermLabel = termName,
                        TrimUnavailable = true
                    };

                    var matchingTerms = termSet.GetTerms(termQuery);
                    clientContext.Load(matchingTerms);
                    clientContext.ExecuteQuery();
                    foreach (var term in matchingTerms)
                    {
                        GuidTerm = term.Id.ToString();
                        break;
                    }
                    //controllo che il termini sia stato trovato
                    if (GuidTerm != "")
                    {
                        break;
                    }
                }
                return GuidTerm;
            }
            catch (Exception e)
            {
                log.LogError("Errore recupero id term group taxonomy, source list: {1}",e.Message);
                //Console.WriteLine("Errore recupero term group " + "// " + e.Message);
                throw new Exception("Errore Recupero id term group: " + e.Message);
            }
        
        }

        public static List<string> getValueTermSet(ClientContext clientContext,string termSetName,ILogger log)
        {
            List<string> allLabels = new List<string>();
            try
            {
                // Get the TaxonomySession
                Tax.TaxonomySession taxonomySession = Tax.TaxonomySession.GetTaxonomySession(clientContext);
                //MMS
                Guid guidStore = new Guid(_MMS);
                Tax.TermStore termStore = taxonomySession.TermStores.GetById(guidStore);
                //Guid Term Group - cellini
                System.Guid guidPeople = new System.Guid(_TermGroup);
                Tax.TermGroup termGroup = termStore.GetGroup(guidPeople);
                //get term set - Sezioni,Building,...
                Tax.TermSetCollection termSetColl = termGroup.TermSets;
                clientContext.Load(termSetColl);
                // Execute the query to the server
                clientContext.ExecuteQuery();
                // Loop through all the termsets
                foreach (Tax.TermSet termSet in termSetColl)
                {
                    if(termSet.Name == termSetName){
                        var terms = termSet.Terms;
                        clientContext.Load(terms);
                        clientContext.ExecuteQuery();
                        foreach (var item in terms)
                        {
                            var label = item.Labels;
                            clientContext.Load(label);
                            clientContext.ExecuteQuery();
                            foreach (var item2 in label)
                            {
                                allLabels.Add(item2.Value.ToString());
                            }
                            
                        }
                        
                    }

                }
                return allLabels;
            }
            catch (Exception e)
            {
                log.LogError("Errore recupero valori term set: {1}",e.Message);
                throw e;
            }
        
        }

        public static string getIdTermSet(ClientContext clientContext,string termSetName,ILogger log)
        {
            var termSetId = "";
            try
            {
                // Get the TaxonomySession
                Tax.TaxonomySession taxonomySession = Tax.TaxonomySession.GetTaxonomySession(clientContext);
                //MMS
                Guid guidStore = new Guid(_MMS);
                Tax.TermStore termStore = taxonomySession.TermStores.GetById(guidStore);
                //Guid Term Group - Unifor
                System.Guid guidPeople = new System.Guid(_TermGroup);
                Tax.TermGroup termGroup = termStore.GetGroup(guidPeople);
                
                Tax.TermSetCollection termSetColl = termGroup.TermSets;
                clientContext.Load(termSetColl);
                // Execute the query to the server
                clientContext.ExecuteQuery();

                
                // Loop through all the termsets
                foreach (Tax.TermSet termSet in termSetColl)
                {
                    if(termSet.Name == termSetName) {
                        termSetId =  termSet.Id.ToString();
                        break;
                    }

                }
                return termSetId;
            }
            catch (Exception e)
            {
                log.LogError("Errore recupero ID term set: {1}",e.Message);
                throw e;
            }
        
        }

        public static void setDefaulValueColumn(ClientContext context,
            string pathFolder,string nameTerm,string valueTerm,ILogger log){

                try
                {
                  //var folder = context.Web.GetFolderByServerRelativeUrl("Shared Documents/General/2020");
                  var folder = context.Web.GetFolderByServerRelativeUrl(_internalNameList +"/" + pathFolder);
                  context.Load(folder);
                  context.ExecuteQuery();

                  SPClient.List list = context.Web.Lists.GetByTitle(_ListLibraryName);

                  //Setting the Metadata Defaults for the Folder
                  MetadataDefaults metadataDefaults = new MetadataDefaults(context, list);
                  //metadataDefaults.SetFieldDefault(folder, "Anno", "2;#" + "2018|0bee8a32-d494-4855-b604-3f4ae78dd980");
                  var prefixTerm = new Random().Next(1, 9) + ";#";
                  metadataDefaults.SetFieldDefault(folder, nameTerm, prefixTerm + valueTerm);
                  metadataDefaults.Update();
                  list.Update();
                  context.ExecuteQuery();
                }
                catch (Exception e )
                {
                    log.LogError("Errore impostazione default column: {1}",e.Message);
                    //Console.WriteLine("Errore set metadata to column " + e.Message);
                    throw new Exception("Errore set metadata to column: " + e.Message);
                }
        }

        public static void CreateFolder(ClientContext context, string siteUrl,string relativePath,string folderName,ILogger log)
        {

            try
            {
                  SPClient.List list = context.Web.Lists.GetByTitle(_ListLibraryName);

                  ListItemCreationInformation newItem = new ListItemCreationInformation();
                  newItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                  newItem.FolderUrl = siteUrl + "/" + _internalNameList;
                  if (!relativePath.Equals(string.Empty))
                  {
                    newItem.FolderUrl += "/" + relativePath;
                  }
                  newItem.LeafName = folderName;
                  SPClient.ListItem item = list.AddItem(newItem);
                  item.Update();
                  context.ExecuteQuery();
            }
            catch (Exception e)
            {
                log.LogError("Errore creazione folder: {1}",e.Message);
              //Console.WriteLine("Errore Creazione folder " + relativePath + " " + folderName + " //" + e.Message);
              throw new Exception("Errore creazione folder: " +e.Message);
            }
    
            
        }




    public static void createContentType(ClientContext context,ILogger log)
    {
      try
      {

        ContentTypeCollection oContentTypeCollection = context.Web.ContentTypes;
        // Load content type collection
        context.Load(oContentTypeCollection);
        context.ExecuteQuery();
        
        ContentType oparentContentType = (from contentType in oContentTypeCollection where contentType.Name == "Document" select contentType).FirstOrDefault();
        ContentTypeCreationInformation oContentTypeCreationInformation = new ContentTypeCreationInformation();
        // Name of the new content type
        oContentTypeCreationInformation.Name = _nameContentType;
        oContentTypeCreationInformation.Group = _GroupColumn;
        oContentTypeCreationInformation.ParentContentType = oparentContentType;
        ContentType oContentType = oContentTypeCollection.Add(oContentTypeCreationInformation);
        context.ExecuteQuery();
      }
      catch (Exception e)
      {
        log.LogError("Errore creazione content-type: {1}",e.Message);
        if(e.HResult == -2146233079){
            throw new Exception("c-404");
        }else{
            throw e;
        }
      }
    }


    public static void createTextOrChoiceColumn(ClientContext context, string internalNameColumn, string displayName, string typeColumn,ILogger log)
    {
        try
        {
            var flow = new JsonObject();
            flow.Schema = "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json";
            flow.elmType = "button";
            flow.customRowAction = new CustomRowAction(){
                action = "executeFlow",
                actionParams = "{\"id\": \"18a46964-ffed-42a5-9e1a-ff3cff633251\"}"
            };
            flow.attributes = new Attributes(){@class = "ms-fontColor-themePrimary ms-fontColor-themeDarker--hover"};
            flow.style = new Style(){border = "none",BackgroundColor = "#456275",cursor = "pointer",display="block"};
            flow.children = new List<Child>(){
                new Child(){
                    elmType = "span",
                    attributes = new Attributes2(){iconName = "Play"},
                    style = new Style2(){color="white",PaddingRight="6px"}
                },
                new Child(){
                    elmType = "span",
                    txtContent = "Pubblica",
                    style = new Style2(){color="white",PaddingRight="6px"}
                }
            };
            var jsonFormat = JsonConvert.SerializeObject(flow);
            string schemaXml = "";
            Web rootWeb = context.Site.RootWeb;
            if (typeColumn == "multiline")
            {
                  schemaXml = @"<Field Type='Note' Name='" + internalNameColumn + "'  DisplayName = '"+displayName+"' NumLines = '10' RichText = 'FALSE' Sortable = 'FALSE' Group='"+ _GroupColumn +"'/> ";
                  
            }
            else if(typeColumn == "choice")
            {
                schemaXml = "<Field Type='Choice' DisplayName='"+displayName+"' Name='"+internalNameColumn+"' Format = 'Dropdown' Group ='"+_GroupColumn+"'>"
                   + "<Default>Non Pubblicato</Default>"
                   + "<CHOICES>"
                   + "    <CHOICE>In Corso</CHOICE>"
                   + "    <CHOICE>Pubblicato</CHOICE>"
                   + "    <CHOICE>Errore</CHOICE>"
                   + "</CHOICES>"
                   + "</Field>";
            }
            else
            {
                schemaXml = @"<Field Type='Text' Name='" + internalNameColumn + "'  DisplayName = '" + displayName + "'  Group='" + _GroupColumn + "' CustomFormatter='"+jsonFormat+"'/> ";
            }
              Field field = rootWeb.Fields.AddFieldAsXml(schemaXml, true, AddFieldOptions.AddFieldToDefaultView);
              context.Load(field);
              context.ExecuteQuery();
              //li aggiungo al content type creato in precedenza
              ContentType sessionContentType = rootWeb.ContentTypes.GetByName(_nameContentType);

              sessionContentType.FieldLinks.Add(new FieldLinkCreationInformation
              {
                Field = field
              });

              sessionContentType.Update(true);
              context.ExecuteQuery();
        }
        catch (System.Exception e)
        {
            log.LogError("Errore creazione colonna: {1}",e.Message);
            throw e;
        }
        
    }

        public static void createColumns(ClientContext context,ILogger log)
        {
            //recupero id del term set
            Web rootWeb = context.Site.RootWeb;
            //li aggiungo al content type creato in precedenza
            ContentType sessionContentType = rootWeb.ContentTypes.GetByName(_nameContentType);
            foreach (KeyValuePair<string, string> pairs in columnContentType)
            {
                var internalName = pairs.Key;
                var displayName = pairs.Value.Split(",")[1];
                var termSetId  = pairs.Value.Split(",")[0];
                try
                {
                  
                  var field = rootWeb.Fields.AddFieldAsXml(@"<Field Type = 'TaxonomyFieldType' Name='" + internalName + "' DisplayName='" + displayName +
                        "' ShowField='Term1033' REquired='FALSE' EnforceUniqueValues='FALSE' Group='" + _GroupColumn + " '/>", true, AddFieldOptions.AddFieldToDefaultView);
                  context.Load(field);
                  //context.ExecuteQuery();

                  Tax.TaxonomyField taxonomyField = context.CastTo<Tax.TaxonomyField>(field);
                  taxonomyField.SspId = new Guid(_MMS);
                  taxonomyField.TermSetId = new Guid(termSetId);
                  taxonomyField.TargetTemplate = String.Empty;
                  taxonomyField.AnchorId = Guid.Empty;
                  if (displayName == "Project Name")
                  {
                    taxonomyField.IsPathRendered = true;
                  }
                  taxonomyField.Update();
                  sessionContentType.FieldLinks.Add(new FieldLinkCreationInformation
                  {
                    Field = field
                  });
                  sessionContentType.Update(true);
            
                }
                catch (Exception e)
                {
                    log.LogError("Errore creazioen colonne {0}",e.Message);
                  //Console.WriteLine("Errore creazione colonna: " + displayName + " " + e.Message);
                    throw new Exception("Errore creazione Colonna: " + e.Message);
                }
            }
            context.ExecuteQuery();
        }


    public static ContentType GetByName(this ContentTypeCollection cts, string name)
        {
            var ctx = cts.Context;
            ctx.Load(cts);
            ctx.ExecuteQuery();
            return Enumerable.FirstOrDefault(cts, ct => ct.Name == name);
        }

        public static void createColumnToSiteColumn(ClientContext context, string internalNameColumn, string displayName,ILogger log)
        {
            string column;
            if (columnContentType.TryGetValue(internalNameColumn, out column))
            {
                  
                  try
                  {
                    //recupero id del term set
                    Web rootWeb = context.Site.RootWeb;
                    var field = rootWeb.Fields.AddFieldAsXml(@"<Field Type = 'TaxonomyFieldType' Name='" + internalNameColumn + "' DisplayName='" + displayName + "' ShowField='Term1033' REquired='FALSE' EnforceUniqueValues='FALSE' Group='" + _GroupColumn + " '/>", true, AddFieldOptions.AddFieldToDefaultView);
                    context.Load(field);
                    //context.ExecuteQuery();

                    Tax.TaxonomyField taxonomyField = context.CastTo<Tax.TaxonomyField>(field);
                    taxonomyField.SspId = new Guid(_MMS);
                    taxonomyField.TermSetId = new Guid(column);
                    taxonomyField.TargetTemplate = String.Empty;
                    taxonomyField.AnchorId = Guid.Empty;
                    if (displayName == "Project Name")
                    {
                      taxonomyField.IsPathRendered = true;
                    }
                    taxonomyField.Update();

                    //li aggiungo al content type creato in precedenza
                    ContentType sessionContentType = rootWeb.ContentTypes.GetByName(_nameContentType);

                    sessionContentType.FieldLinks.Add(new FieldLinkCreationInformation
                    {
                      Field = field
                    });

                    sessionContentType.Update(true);
                    context.ExecuteQuery();
                  }
                  catch (Exception e)
                  {
                      log.LogError("Errore creazione colonna a livello di site column: {0}",e.Message);
                      //Console.WriteLine("Errore creazione colonna: " + displayName + " " + e.Message);
                      throw new Exception("Errore creazione Colonna: " + e.Message) ;
                  }
            }
        }


       

          public static void setContentTypeToList(ClientContext clientContext,ILogger log){
            try
            {
                ContentTypeCollection contentTypeCollection;
                // Option - 1 - Get Content Types from Root web
                //contentTypeCollection = clientContext.Site.RootWeb.ContentTypes;
                
                contentTypeCollection = clientContext.Web.ContentTypes;
                
                clientContext.Load(contentTypeCollection);
                clientContext.ExecuteQuery();
                
                ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == _nameContentType select contentType).FirstOrDefault();
                
                List targetList = clientContext.Web.Lists.GetByTitle(_ListLibraryName);
                targetList.ContentTypes.AddExistingContentType(targetContentType);
                targetList.Update();
                clientContext.Web.Update();
                clientContext.ExecuteQuery();

                ContentTypeCollection currentCtOrder = targetList.ContentTypes;
                clientContext.Load(currentCtOrder);
                clientContext.ExecuteQuery();

                IList<ContentTypeId> reverceOrder = new List<ContentTypeId>();
                foreach (ContentType ct in currentCtOrder)
                {
                    if (ct.Name.Equals(_nameContentType))
                    {
                        reverceOrder.Add(ct.Id);
                    }
                }
                targetList.ContentTypesEnabled = true;
                targetList.RootFolder.UniqueContentTypeOrder = reverceOrder;
                targetList.RootFolder.Update();
                targetList.Update();
                clientContext.ExecuteQuery();
            }
            catch (System.Exception e)
            {
                log.LogError("Errore impostazione content-type  a livello di lista: {0}",e.Message);
                //Console.WriteLine("Errore aggiunta content type alla document Library- " + e.Message);
                throw new Exception("Errore set Content Type to list: " + e.Message);
            }
            
        }

        public static void setViewList(ClientContext clientContext, string checkTeam,ILogger log){
            try
            {
                List targetList = clientContext.Web.Lists.GetByTitle(_ListLibraryName);
                // Get required view by specifying view Title here
                View targetView = targetList.Views.GetByTitle("All Documents");
                targetView.DefaultView = false;
                
                // Set the view as default
                targetView.DefaultView = false;
                targetView.ViewFields.Add("Canale");
                targetView.ViewFields.Add("Project Name");
                targetView.ViewFields.Add("PM");
                targetView.ViewFields.Add("Anno");
                targetView.ViewFields.Add("ISO");
                targetView.ViewFields.Add("Citta");
                targetView.ViewFields.Add("Numero Offerta");
                targetView.ViewFields.Add("Tipo Documento");
                if(checkTeam == "Team"){
                    targetView.ViewFields.Add("Commenti Pubblicazione");
                }else{
                    targetView.ViewFields.Add("Stato");
                }

                targetView.Update();
                clientContext.ExecuteQuery();
            }
            catch (System.Exception e)
            {
                log.LogError("Errore impostazione della vista: {0}",e.Message);
                //Console.WriteLine("Errore impostazione vista di default " + e.Message);
                throw new Exception("Errore set view: " + e.Message);
            }            
        }

        public static void createViewList(ClientContext clientContext, string nameView,string typeChannel,ILogger log){
            try
            {
                List targetList = clientContext.Web.Lists.GetByTitle(_ListLibraryName);
                // Get required view by specifying view Title here
                ViewCollection viewCollection = targetList.Views;
                clientContext.Load(viewCollection);
                ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
                viewCreationInformation.Title = nameView;
                viewCreationInformation.SetAsDefaultView = true;
                string CommaSeparateColumnNames = "";
                if(typeChannel == "Team"){
                    CommaSeparateColumnNames = "Name,Canale,Project Name,PM,Anno,ISO,Citta,Numero Offerta,Tipo Documento,Commenti Pubblicazione";
                }else{
                    CommaSeparateColumnNames = "Name,Canale,Project Name,PM,Anno,ISO,Citta,Numero Offerta,Tipo Documento,Stato,Pubblica";
                }
                
                viewCreationInformation.ViewFields = CommaSeparateColumnNames.Split(',');
                View listView = viewCollection.Add(viewCreationInformation);
                //clientContext.ExecuteQuery();
                
                listView.Scope = ViewScope.Recursive;
                listView.Update();
                clientContext.ExecuteQuery();
            }
            catch (System.Exception e)
            {
                log.LogError("Errore creazione della vista: {0}",e.Message);
                //Console.WriteLine("Errore creazione della vista " + e.Message);
                throw new Exception("Errore create view: " +e.Message);
            }
            

        }




    }
}
