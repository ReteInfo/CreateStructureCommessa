using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using graph = Microsoft.Graph;
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
using Microsoft.Extensions.Logging;

namespace Company.Function
{
    public static class Service
    {

        //private static string _UrlHubSite = Environment.GetEnvironmentVariable("urlHubSite");
        private static string _UrlHubSite = "https://uniforspa.sharepoint.com/sites/Commesse";
        //private static string _urlFunctionStructure = Environment.GetEnvironmentVariable("EndpointFunction");
        private static string _urlFunctionStructure = "http://localhost:7071/api/CreateTeamStructure?name=";
        private static string _configListCommesseCreate = "Elenco Commesse Create";
        private static string _configList = "Configurations";

        public static void insertObjectToListConfig(string token,TeamCommessa objectTeam, ILogger log){
            try
            {
                using (ClientContext clientContext = new ClientContext(_UrlHubSite))
                {
                    clientContext.ExecutingWebRequest += (sender, e) =>
                    {
                        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
                    };
                    List targetList = clientContext.Web.Lists.GetByTitle(_configListCommesseCreate);
                    
                    ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
                    ListItem oItem = targetList.AddItem(oListItemCreationInformation);
                    oItem["Title"] = Regex.Replace(objectTeam.NameTeam, @" ", "");
                    oItem["ConfigurationObject"] = JsonConvert.SerializeObject(objectTeam);
                    oItem["CreazioneStrutture"] = _urlFunctionStructure+Regex.Replace(objectTeam.NameTeam, @" ", "");
                    oItem["Stato"] = "Strutture da creare";
                    
                    oItem.Update();
                    clientContext.ExecuteQuery();
                }
            }
            catch (System.Exception e)
            {
                log.LogError("Errore inserimento oggetto commessa nell'elenco commesse - Elenco commesse create: {0}",e.Message);
                throw e;
            }
        }

        public static string getValueConfigFile(string nameJson,string token, ILogger log)
        {
            try
            {
                var  config = "";
            
                using(var clientContext = new ClientContext(_UrlHubSite)){
                    clientContext.ExecutingWebRequest += (sender, e) =>
                    {
                        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
                    };   


                    SPClient.List oList = clientContext.Web.Lists.GetByTitle(_configList);

                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                        "<Value Type='Text'>"+nameJson+"</Value></Eq></Where></Query></View>";

                    ListItemCollection collListItem = oList.GetItems(camlQuery);

                    clientContext.Load(collListItem);

                    clientContext.ExecuteQuery();

                    foreach (var item in collListItem)
                    {
                        config = item["Value"].ToString();
                    }

                }
                return config;
            }
            catch (System.Exception e)
            {
                log.LogError("Errore recupero oggetto dalla lista - Configurations : {0}",e.Message);
                throw e;
            }
        }

    public static void updateObjectToListConfig(string token,TeamCommessa objectTeam,string message,string stato,ILogger log){
        try
        {
            var note = "";
            using (ClientContext clientContext = new ClientContext(_UrlHubSite))
            {
                clientContext.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
                };
                List targetList = clientContext.Web.Lists.GetByTitle(_configListCommesseCreate);
                
                CamlQuery oQuery = new CamlQuery();
                    oQuery.ViewXml = String.Format(@"<View><Query><Where>
                    <Eq>
                    <FieldRef Name='Title' />
                    <Value Type='Text'>{0}</Value>
                    </Eq>
                    </Where></Query></View>",Regex.Replace(objectTeam.NameTeam, @" ", ""));
                
                ListItemCollection oItems = targetList.GetItems(oQuery);
                clientContext.Load(oItems);
                clientContext.ExecuteQuery();
                
                foreach (var item in oItems)
                {
                    if(stato != ""){
                        item["Stato"] = stato;
                        if(stato == "Errore" || stato == "Da completare"){
                            item["ConfigurationObject"] = JsonConvert.SerializeObject(objectTeam);
                        }
                    }
                    if(message == "Strutture già create"){
                        item["Note"] = message;
                    }else{
                        if(item["Note"] != null){
                            note = item["Note"].ToString();
                        }
                    }
                    
                    item["Note"] = note + " " + message;
                    item.Update();
                }
                
                clientContext.ExecuteQuery();
            }
        }
        catch (System.Exception e)
        {
            log.LogError("Errore aggiornamento oggetto commessa nell'elenco commesse: {0}",e.Message);
            //Console.WriteLine(e.Message + " " + e.StackTrace);
            throw e;
        }
    }

    public static TeamCommessa getObjectToListConfig(string token,string nameTeam,ILogger log){
        try
        {
            var stato = "";
            var objectString = "";
            
            using (ClientContext clientContext = new ClientContext(_UrlHubSite))
            {
                clientContext.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
                };
                List targetList = clientContext.Web.Lists.GetByTitle(_configListCommesseCreate);
                
                CamlQuery oQuery = new CamlQuery();
                    oQuery.ViewXml = String.Format(@"<View><Query><Where>
                    <Eq>
                    <FieldRef Name='Title' />
                    <Value Type='Text'>{0}</Value>
                    </Eq>
                    </Where></Query></View>",nameTeam);
                
                ListItemCollection oItems = targetList.GetItems(oQuery);
                clientContext.Load(oItems);
                clientContext.ExecuteQuery();
                if(oItems.Count == 0){
                    throw new Exception("Nessuna commessa trovata!");
                }else{
                    foreach (var item in oItems)
                    {
                        objectString = item["ConfigurationObject"].ToString();
                        stato = item["Stato"].ToString();
                    }
                }
            }

            TeamCommessa objectTeam = JsonConvert.DeserializeObject<TeamCommessa>(objectString);
            objectTeam.StatoCreazione = stato;
            return objectTeam;
        }
        catch (System.Exception e)
        {
            log.LogError("Errore recupero oggetto commessa nell'elenco commesse: {0}",e.Message);
            //Console.WriteLine(e.Message + " " + e.StackTrace);
            throw e;
        }
    }
    
        public static async Task<List<string>> getIdUserFromEmail(graph.GraphServiceClient graphClient,List<string> emails,ILogger log){
            List<string> guids = new List<string>();
            var email = "";
            try
            {
                foreach (var item in emails)
                {
                  try
                  {
                    email = item;
                    var result = await graphClient.Users.Request()
                        .Filter(String.Format("mail eq '{0}'", item))
                        .GetAsync();
                    guids.Add(result[0].Id);
                  }
                  catch (System.Exception e)
                  {
                    //Console.WriteLine("Errore recupero GUID user con email: " + email + " errore " + e.Message);
                    continue;
                  }

                }
            }
            catch (Exception e)
            {
                log.LogError("Errore recupero id utente dall'email: {0} : {1}",email, e.Message);    
                throw e;
            }
            
            return guids;
        }

        public async static Task<graph.Group> createGroup(graph.GraphServiceClient graphServiceClient, string nameTeam, string nameTeamEmail, ILogger log)
        {
            try
            {
                    var group = new graph.Group
                {
                    Description = "testalpha",
                    DisplayName = nameTeam,
                    GroupTypes = new List<String>()
                    {
                        "Unified"
                    },
                    MailEnabled = true,
                    MailNickname = nameTeamEmail,
                    SecurityEnabled = false,
                    Visibility = "Private"

                };

                return await graphServiceClient.Groups
                    .Request()
                    .AddAsync(group);
            }
            catch (System.Exception e)
            {
                //Console.WriteLine("Errore creazione gruppo " + e.Message);
                log.LogError("ERRORE creazione gruppo: {0}",e.Message);
                throw e;
            }

        }

        public async static Task<graph.Team> createTeam(graph.GraphServiceClient graphServiceClient, string idGroup,string nameTeam,ILogger log)
        {
            try
            {
                var team = new graph.Team
                {
                    DisplayName = nameTeam,
                    ODataType = null
                };


                return await graphServiceClient.Groups[idGroup].Team
                    .Request()
                    .PutAsync(team);
            }
            catch (System.Exception e)
            {
                log.LogError("ERRORE creazione team:{0}",e.Message);
                //Console.WriteLine("Errore creazione team " + e.Message);
                throw e;
            }
            
        }

        public static  void associateToHubSite(string urlSiteToAssociate,ILogger log){
        try
        {
            var token = Authentication.authHubSite(log).Result;
            using (var clientContext = new ClientContext(Authentication.resourceAdmin))
            {
                clientContext.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
                };
                var tenant = new Tenant(clientContext);   

                tenant.ConnectSiteToHubSite(urlSiteToAssociate, _UrlHubSite);
                clientContext.ExecuteQuery();
                log.LogInformation("Sito {0} associato all'hub site",urlSiteToAssociate);
            }
        }
        catch (System.Exception e)
        {
            log.LogInformation("Errore associazione sito all'Hub site: {1}",e.Message);
            throw new Exception("Errore associazione all' Hub site");
        }
    }

        public async static Task<graph.Channel> createChannel(graph.GraphServiceClient graphServiceClient,string teamId,string ownerId,string nameChannel,ILogger log)
        {

            try
            {
                var channel = new graph.Channel
                {
                    MembershipType = graph.ChannelMembershipType.Private,
                    DisplayName = nameChannel,
                    Description = "This is my first private channels",
                    Members = new graph.ChannelMembersCollectionPage()
                    {
                        new graph.AadUserConversationMember
                        {
                            Roles = new List<string> { "owner" },
                            //UserId = ownerId
                            AdditionalData = new Dictionary<string, object>
                            {
                                ["user@odata.bind"] = $"https://graph.microsoft.com/beta/users('{ownerId}')"
                            }
                        }
                    }
                };
                graph.Channel channels  = await graphServiceClient.Teams[teamId].Channels.Request().AddAsync(channel);
                return channels;

            }
            catch (System.Exception e)
            {
                log.LogError("Errore creazione canale {0}: {1}",nameChannel,e.Message);
                //Console.WriteLine("Errore creazione canale " + nameChannel + " " + e.Message);
                throw e;
            }
        }

        //GESTIONE DEI MEMBRI E OWNER
        public async static Task addMoreOwnerToChannel(graph.GraphServiceClient graphServiceClient,string teamId,string channelId,List<string> owners,ILogger log)
        {          
            var owner = "";  
            foreach (var item2 in owners)
            {
                owner = item2;
                try
                {
                    var conversationMember = new graph.AadUserConversationMember
                    {
                        Roles = new List<String>(){"owner"},
                        AdditionalData = new Dictionary<string, object>()
                            {
                                ["user@odata.bind"] = $"https://graph.microsoft.com/beta/users('{item2}')"
                            }
                    };

                    await graphServiceClient.Teams[teamId].Channels[channelId].Members
                        .Request()
                        .AddAsync(conversationMember);
                }
                catch (System.Exception e)
                {
                    log.LogError("Errore aggiunta owner {0} al canale {1}: {2}",owner,channelId,e.Message);
                    //Console.WriteLine("Errore aggiunta owner al canale " + e.Message);
                    continue;
                }
                
            }
        }
        public async static Task addMemberToChannel(graph.GraphServiceClient graphServiceClient, string teamId, string channelId,List<string> memberId,ILogger log)
        {
            var member = "";
            foreach (var item2 in memberId)
            {
                try
                {
                    member = item2;
                    var conversationMember = new graph.AadUserConversationMember
                    {
                        Roles = new List<String>()
                        {
                        },
                        AdditionalData = new Dictionary<string, object>()
                            {
                                //$"https://graph.microsoft.com/beta/users('{item2}')"
                                ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users/"+item2
                            }
                    };

                    await graphServiceClient.Teams[teamId].Channels[channelId].Members
                        .Request()
                        .AddAsync(conversationMember);
                }
                catch (System.Exception e)
                {
                    log.LogError("Errore aggiunta member {0} al canale {1}: {2}",member,channelId,e.Message);
                    //Console.WriteLine("Errore aggiunta membro al canale" + e.Message);
                    throw e;
                }  
            }
        }
        public async static Task addOwnerToGroup(graph.GraphServiceClient graphServiceClient,string idGroup, List<string> idOwners, ILogger log)
        {
            var owner = "";
            foreach (var item in idOwners)
            {
                try
                {
                    owner = item;
                    var directoryObject = new graph.DirectoryObject
                    {
                        Id = item
                    };
                    //id del gruppo creato
                    await graphServiceClient.Groups[idGroup].Owners.References
                        .Request()
                        .AddAsync(directoryObject);
                }
                catch (System.Exception e)
                {
                    log.LogError("Errore aggiunta owner {0} al gruppo : {1}",owner, e.Message);
                    //Console.WriteLine("Errore aggiunta owner al gruppo " + e.Message);
                    continue;
                }
                
            }

        }

        public async static Task addMemberToGroup(graph.GraphServiceClient graphServiceClient, string idGroup, List<string> idMembers,ILogger log)
        {
            var member = "";
            foreach (var item in idMembers)
            {
                try{
                    member = item;
                    var directoryObject = new graph.DirectoryObject
                    {
                        Id = item
                    };

                    await graphServiceClient.Groups[idGroup].Members.References
                        .Request()
                        .AddAsync(directoryObject);

                }catch(Exception e){
                    log.LogError("ERRORE aggiunta membro {0} al gruppo: {1}",member,e.Message);
                    //Console.WriteLine("Errore aggiunta membro al gruppo " + e.Message);
                    continue;
                }
                
            }
            
        }
        
    }
}
