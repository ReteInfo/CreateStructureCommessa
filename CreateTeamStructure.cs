using System;
using System.IO;
using System.Net;
using generic = System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using System.Threading;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Linq;
using OfficeDevPnP.Core.Pages;



namespace Company.Function
{
    public static class CreateTeamStructure
    {
        private static ConfigCanali usersDefault;
        private static ConfigFolder jsonConfig;
        private static string siteUrl = Documents.siteUrl;


        [FunctionName("CreateTeamStructure")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            try
            {
                Result allOperation = new Result();
                //scelgo il tipo di operazione
                string name = req.Query["name"];
                if (name == "" || name == null)
                {
                    log.LogError("Errore chiamata:{0} - nessun paramentro riscontrato",req);
                    throw new Exception("Errore richiesta");
                }

                //recupero l'oggetto
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                var objectTeam = JsonConvert.DeserializeObject<TeamCommessa>(requestBody);

                switch (name)
                {
                    case "CreateTeam":
                        //inserisco l'oggetto nella lista per creare successivamente le DL
                        createObjectToListConfig(objectTeam,log);
                        var group = await createStructureTeam(objectTeam,log);
                        Service.associateToHubSite(siteUrl + Regex.Replace(objectTeam.NameTeam, @" ", ""),log);
                        //ritorno l'oggetto
                        allOperation.IdGroup = group.Id;
                        allOperation.Operation = "OK!";
                        break;
                    case "CreateChannel":
                        await createChannel(objectTeam,log);
                        allOperation.Operation = "OK!";
                        break;
                    default:
                        //creazione delle strutture
                        //Console.WriteLine("---------------------Chiamata Operazione struttura document library sul team " + name);
                        var token = Authentication.authSP(log).Result;
                        TeamCommessa obj = Service.getObjectToListConfig(token,name,log);
                        if(obj.StatoCreazione == "Strutture da creare" || obj.StatoCreazione == "Da completare"){
                            try{
                                initializeStructureDL(obj,token,log);
                                allOperation.Operation = "Creazione strutture avvenuta con successo";
                            }catch(Exception e){
                                allOperation.Operation = e.Message;
                            }   
                        }else{
                            if(obj.StatoCreazione == "Errore"){
                                Service.updateObjectToListConfig(token,obj,"Errore creazione strutture",obj.StatoCreazione,log);
                                allOperation.Operation = "Errore...";
                            }else{
                                Service.updateObjectToListConfig(token,obj,"Strutture già create",obj.StatoCreazione,log);
                                allOperation.Operation = "Strutture già create per tutti i canali";
                            }
                            
                        }
                        return new OkObjectResult(allOperation.Operation);
                }

                return new OkObjectResult(allOperation);
            }
            catch (System.Exception e)
            {
                var result = new ObjectResult(e.Message);
                if (e.Message == "Errore richiesta")
                {
                    result.StatusCode = StatusCodes.Status400BadRequest;
                }
                return result;
            }
        }


        public static async Task<GroupTeam> createStructureTeam(TeamCommessa teamObject, ILogger log)
        {
            try
            {
                GraphServiceClient graphClient = await Authentication.auth(log);
                var allOwners = removeDuplicateFromList(teamObject.ownersTeam);

                var resultGuidsOwner = await Service.getIdUserFromEmail(graphClient, allOwners,log);

                var token = Authentication.authSP(log).Result;
                usersDefault = JsonConvert.DeserializeObject<ConfigCanali>(Service.getValueConfigFile("ConfigUserChannels", token,log));
                //recupero ID member channel
                generic.List<string> resultIdChannelGuids;

                //recupero tutti i membri di default
                var membersDefault = usersDefault.getAllMembers();
                var allMembers = membersDefault.Union(teamObject.membersChannel).ToList();
                resultIdChannelGuids = await Service.getIdUserFromEmail(graphClient, allMembers,log);


                string _NameTeamEmail = Regex.Replace(teamObject.NameTeamEmail, @" ", "");
                //creo il gruppo
                var resultGroup = await Service.createGroup(graphClient, teamObject.NameTeam, _NameTeamEmail, log);
                log.LogInformation("Creazione gruppo - {0} - avvenuta con successo",teamObject.NameTeam);
                Thread.Sleep(8000);

                //aggiungo gli owners al gruppo
                await Service.addOwnerToGroup(graphClient, resultGroup.Id, resultGuidsOwner,log);
                log.LogInformation("Aggiunta Owner al  gruppo - {0} - avvenuta con successo",teamObject.NameTeam);
                Thread.Sleep(3000);

                //aggiungo i membri al gruppo
                await Service.addMemberToGroup(graphClient, resultGroup.Id, resultIdChannelGuids, log);
                log.LogInformation("Aggiunta Members al  gruppo - {0} - avvenuta con successo",teamObject.NameTeam);
                Thread.Sleep(5000);
                
                //creo il team - quando creo il gruppo provato in automatico viene creato il team
                var resultTeam = await Service.createTeam(graphClient, resultGroup.Id, teamObject.NameTeam,log);
                log.LogInformation("Creazione team - {0} - avvenuta con successo",teamObject.NameTeam);

                var group = new GroupTeam();
                group.Id = resultGroup.Id;

                return group;

            }
            catch (System.Exception e)
            {
                throw e;
            }
        }

        public static async Task createChannel(TeamCommessa teamObject, ILogger log)
        {
            try
            {
                GraphServiceClient graphClient = await Authentication.auth(log);
                //recupero i valori di default dei membri
                var token = Authentication.authSP(log).Result;
                usersDefault = JsonConvert.DeserializeObject<ConfigCanali>(Service.getValueConfigFile("ConfigUserChannels", token,log));
                log.LogInformation("Recupero oggetto di configurazione membri canali avvenuta con successo");

                var allOwners = removeDuplicateFromList(teamObject.ownersTeam);
                //recupero ID owner Team tramite email - owner team e canali soo gli stessi
                var resultGuidsOwner = await Service.getIdUserFromEmail(graphClient, allOwners,log);

                //creo i canali con all'interno almeno un owner
                foreach (var item in teamObject.Channels)
                {
                    try
                    {
                        var channel = await Service.createChannel(graphClient, teamObject.IdGroup, resultGuidsOwner[0], item.NameChannel,log);
                        item.IdChannel = channel.Id;
                        log.LogInformation("Creazione canale {0} avvenuta con successo",item.NameChannel);
                        //Console.WriteLine("--------Creazione canale " + item.NameChannel + " avvenuta con successo!");
                        Thread.Sleep(3000);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
                Thread.Sleep(3000);
                if (allOwners.Count > 1)
                {
                    //aggiungo i restanti owner ai canali
                    foreach (var item in teamObject.Channels)
                    {
                        await Service.addMoreOwnerToChannel(graphClient, teamObject.IdGroup, item.IdChannel, resultGuidsOwner,log);
                    }
                }
                foreach (var item in teamObject.Channels)
                {
                    try
                    {

                        //recupero i membri di default del canale
                        var membersDefault = usersDefault.getMembersChannel(item.NameChannel);
                        //recupero i membri aggiunti
                        var allmembers = membersDefault.Union(item.Members).ToList();
                        var allDistinctmembers = await Service.getIdUserFromEmail(graphClient, allmembers,log);
                        await Service.addMemberToChannel(graphClient, teamObject.IdGroup, item.IdChannel, allDistinctmembers,log);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

            }
            catch (System.Exception e)
            {
                throw e;
            }
        }

        public static void createObjectToListConfig(TeamCommessa objectTeam,ILogger log){
            try
            {
                var token = Authentication.authSP(log).Result;
                Service.insertObjectToListConfig(token,objectTeam,log);
                log.LogInformation("Oggetto di configurazione inserito nell'elenco delle commesse");
            }
            catch (System.Exception e)
            {
                throw new Exception("Errore Inserimento commessa nella lista di configurazione");
            }
        }

        public static void initializeStructureDL(TeamCommessa teamObject, string token, ILogger log)
        {
            var check = false;
            var UrlSite = "";
            try
            {
                //il canale general è l'unico che esiste sempre
                if(teamObject.StatoCreazione == "Strutture da creare"){
                    Service.updateObjectToListConfig(token,teamObject,"","In corso",log);
                    //recupero i valori di default 
                    jsonConfig = JsonConvert.DeserializeObject<ConfigFolder>(Service.getValueConfigFile("ConfigFile", token,log));
                    //imposto i content type,colonne e viste sul canale general
                    UrlSite = siteUrl + Regex.Replace(teamObject.NameTeam, @" ", "");
                    SettingsMajorUpdate(token, UrlSite, "Team",log);
                    log.LogInformation("Creazione colonne,content-type, viste sul canale GENERAL avvenuta con successo");
                    CreateDocumentLibraryStructure(token, UrlSite, teamObject, "Team",log);
                    log.LogInformation("Creazione struttura documentale sul canale GENERAL avvenuta con successo");
                    //Console.WriteLine("Fine creazione struttura folder general");
                    Service.updateObjectToListConfig(token,teamObject,"Canale: General - OK","",log);
                }

                foreach (var channel in teamObject.Channels)
                {
                    try
                    {
                        if(channel.create != true){
                            UrlSite = siteUrl + Regex.Replace(teamObject.NameTeam, @" ", "") + "-" + Regex.Replace(channel.NameChannel, @" ", "");

                            SettingsMajorUpdate(token, UrlSite, "channel",log);
                            log.LogInformation("Creazione colonne,content-type, viste sul canale {0} avvenuta con successo",channel.NameChannel);
                            //Console.WriteLine("Fine creazione struttura folder canale - " + channel.NameChannel);
                            CreateDocumentLibraryStructure(token, UrlSite, teamObject, channel.NameChannel,log);
                            log.LogInformation("Creazione struttura documentale sul canale {0} avvenuta con successo",channel.NameChannel);
                            //Console.WriteLine("Fine impostazioni metadati canale- " + channel.NameChannel);
                            channel.create = true;
                            if(teamObject.StatoCreazione == "Da completare"){
                                Service.updateObjectToListConfig(token,teamObject,"***Canale:"+channel.NameChannel+"- OK",teamObject.StatoCreazione,log);
                            }else{
                                Service.updateObjectToListConfig(token,teamObject,"Canale:"+channel.NameChannel+"- OK",teamObject.StatoCreazione,log);
                            }
                        }
                    }
                    catch (Exception e)
                    {

                        if (e.Message == "Errore creazione Colonna")
                        {
                            throw e;
                        }
                        else if (e.Message == "Errore create view")
                        {
                            throw e;
                        }
                        else if(e.Message == "c-404"){
                            //inserisco l'errore nel canale
                            Service.updateObjectToListConfig(token,teamObject,"Canale:"+channel.NameChannel+"- struttura non creata","",log);
                            check = true;
                            continue;
                        }
                        else
                        {
                            continue;
                        }

                    }
                }
                //faccio un update sullo stato 
                if(check == true){
                    //mancano delle strutture su dei canali
                    Service.updateObjectToListConfig(token,teamObject,"","Da completare",log);
                }else{
                    Service.updateObjectToListConfig(token,teamObject,"Strutture create con successo","Strutture create",log);
                }
            }
            catch (System.Exception e)
            {
                Service.updateObjectToListConfig(token,teamObject,e.Message,"Errore",log);
                throw e;
            }

        }

        public static void SettingsMajorUpdate(string token, string UrlSite,string checkTeam,ILogger log)
        {
            //tempo esecuzione metodo

            using (var context = new ClientContext(UrlSite))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
                };
                try
                {

                    //creo il content-type
                    Documents.createContentType(context, log);
                    //creo le colonne nel site column e le aggiungo al content type
                    Documents.createColumns(context,log);
                    if(checkTeam == "Team")//aggiungo colonna commento pubblicazione
                    {
                        Documents.createTextOrChoiceColumn(context, "CommentiPubblicazione", "Commenti Pubblicazione", "multiline",log);
                    }
                    else
                    {
                        Documents.createTextOrChoiceColumn(context, "Stato", "Stato","choice",log);
                        Documents.createTextOrChoiceColumn(context,"Pubblica", "Pubblica", "flow",log);
                    }
                    //imposto il content type
                    Documents.setContentTypeToList(context,log);
                    //imposto le viste
                    Documents.setViewList(context,checkTeam,log);
                    Documents.createViewList(context, "Teams",checkTeam,log);
                }
                catch (Exception e)
                {
                    throw e;
                }

            }
        }

        public static void CreateDocumentLibraryStructure(string token, string UrlSite, TeamCommessa teamObject, string channel,ILogger log)
        {

            var relativePath = channel;
            if (channel == "Team")
            {
                relativePath = "General";
            }

            using (var context = new ClientContext(UrlSite))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
                };

                foreach (var item in jsonConfig.LevelDescriptors)
                {
                    try
                    {
                        //se non è una foglia
                        if (!item.Leaves)
                        {
                            //controllo se è un metadato
                            var nameFolder = "";
                            //recupero il nome della folder
                            foreach (var valueMetadato in teamObject.Metadata)
                            {
                                var nameMetadato = valueMetadato.NameMetadato;
                                if (nameMetadato == item.Name)
                                {
                                    var valMeta = valueMetadato.ValueMetadato;

                                    if (item.Name == "Customers")
                                    {
                                        valMeta = valueMetadato.ValueMetadato.Split(";")[1];
                                    }
                                    nameFolder = valMeta;
                                    break;
                                }
                            }
                            //creo la folder
                            Documents.CreateFolder(context, UrlSite, relativePath, nameFolder,log);
                            //imposto i metadata
                            if (item.Name == "Years")
                            {
                                var idValueMetadato = Documents.getTaxonomyTermGroup(context, nameFolder,log);
                                if (idValueMetadato != null || idValueMetadato != "")
                                {
                                    Documents.setDefaulValueColumn(context, relativePath + "/" + nameFolder, item.Name + "MT", nameFolder + "|" + idValueMetadato,log);
                                }
                            }
                            else if (item.Name == "Customers")
                            {

                                var idValueMetadato = Documents.getTaxonomyTermGroup(context, nameFolder,log);
                                if (idValueMetadato != null || idValueMetadato != "")
                                {
                                    Documents.setDefaulValueColumn(context, relativePath + "/" + nameFolder, item.Name + "MT", nameFolder + "|" + idValueMetadato,log);
                                }

                                //imposto PM-canale
                                generic.List<string> allSettingCustomers = new generic.List<string>() { "PMs", "Companies", "ISO", "Cities" };
                                foreach (var val in teamObject.Metadata)
                                {
                                    if (allSettingCustomers.Contains(val.NameMetadato))
                                    {
                                        idValueMetadato = Documents.getTaxonomyTermGroup(context, val.ValueMetadato,log);
                                        if (idValueMetadato != null || idValueMetadato != "")
                                        {
                                            Documents.setDefaulValueColumn(context, relativePath + "/" + nameFolder, val.NameMetadato + "MT", val.ValueMetadato + "|" + idValueMetadato,log);
                                        }

                                    }
                                }
                            }
                            else if (item.Name == "SubSuddivisione")
                            {
                                //imposto Number Offer
                                foreach (var val in teamObject.Metadata)
                                {
                                    if (val.NameMetadato == "NumberOffer")
                                    {

                                        var idValueMetadato = Documents.getTaxonomyTermGroup(context, val.ValueMetadato,log);
                                        if (idValueMetadato != null || idValueMetadato != "")
                                        {
                                            Documents.setDefaulValueColumn(context, relativePath + "/" + nameFolder, val.NameMetadato + "MT", val.ValueMetadato + "|" + idValueMetadato,log);
                                            break;
                                        }
                                    }
                                }
                            }
                            relativePath = relativePath + @"\" + nameFolder;
                        }
                        else
                        {
                            if (item.Name == "Doc Types")
                            {
                                generic.List<string> labels = Documents.getValueTermSet(context, item.Name,log);
                                foreach (var label in labels)
                                {
                                    Documents.CreateFolder(context, UrlSite, relativePath, label,log);
                                    var idValueMetadato = Documents.getTaxonomyTermGroup(context, label,log);
                                    if (idValueMetadato != null || idValueMetadato != "")
                                    {
                                        Documents.setDefaulValueColumn(context, relativePath + "/" + label, "DocTypesMT", label + "|" + idValueMetadato,log);
                                    }
                                }
                            }

                        }
                    }
                    catch (Exception e)
                    {
                        if (e.Message == "Errore creazione folder")
                        {
                            throw e;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }

            }

        }

        public static generic.List<string> removeDuplicateFromList(generic.List<string> values)
        {
            generic.List<string> distinctValues = new generic.List<string>();
            foreach (var item in values)
            {
                if (!distinctValues.Contains(item))
                {
                    distinctValues.Add(item);
                }
            }
            return distinctValues;
        }





    }
}
