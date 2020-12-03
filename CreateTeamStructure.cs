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
        private static string siteUrl = "https://reteinformatica.sharepoint.com/sites/";


        [FunctionName("CreateTeamStructure")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            try
            {
                Result allOperation = new Result();
                //scelgo il tipo di operazione
                string name = req.Query["name"];
                if (name == "" || name == null)
                {
                    throw new Exception("Errore richiesta");
                }

                //recupero l'oggetto
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                var objectTeam = JsonConvert.DeserializeObject<TeamCommessa>(requestBody);

                switch (name)
                {
                    case "CreateTeam":
                        Console.WriteLine("---------------------Chiamata Operazione creazione team");
                        //inserisco l'oggetto nella lista per creare successivamente le DL
                        createObjectToListConfig(objectTeam);
                        Console.WriteLine("----oggetto commessa inserita nella lista di configurazione");
                        var group = await createStructureTeam(objectTeam);
                        Console.WriteLine("----Creazione del team avvenuta con successo");
                        Service.associateToHubSite(siteUrl + Regex.Replace(objectTeam.NameTeam, @" ", ""));
                        Console.WriteLine("----sito associato all'HUB site commesse");
                        //ritorno l'oggetto
                        allOperation.IdGroup = group.Id;
                        allOperation.Operation = "OK!";
                        break;
                    case "CreateChannel":
                        Console.WriteLine("---------------------Chiamata Operazione creazione canali");
                        await createChannel(objectTeam);
                        Console.WriteLine("---------- creazione canali avvenuta con successo");
                        allOperation.Operation = "OK!";
                        break;
                    default:
                        //creazione delle strutture
                        Console.WriteLine("---------------------Chiamata Operazione struttura document library sul team " + name);
                        var token = Authentication.authSP().Result;
                        TeamCommessa obj = Service.getObjectToListConfig(token,name);
                        if(obj.StatoCreazione == "Strutture da creare" || obj.StatoCreazione == "Da completare"){
                            try{
                                initializeStructureDL(obj,token);
                                allOperation.Operation = "Creazione strutture avvenuta con successo";
                            }catch(Exception e){
                                allOperation.Operation = e.Message;
                            }   
                        }else{
                            if(obj.StatoCreazione == "Errore"){
                                Service.updateObjectToListConfig(token,obj,"Errore creazione strutture",obj.StatoCreazione);
                                allOperation.Operation = "Errore...";
                            }else{
                                Service.updateObjectToListConfig(token,obj,"Strutture già create",obj.StatoCreazione);
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


        public static async Task<GroupTeam> createStructureTeam(TeamCommessa teamObject)
        {
            try
            {
                GraphServiceClient graphClient = await Authentication.auth();
                var allOwners = removeDuplicateFromList(teamObject.ownersTeam);

                var resultGuidsOwner = await Service.getIdUserFromEmail(graphClient, allOwners);

                var token = Authentication.authSP().Result;
                usersDefault = JsonConvert.DeserializeObject<ConfigCanali>(Service.getValueConfigFile("ConfigUserChannels", token));
                //recupero ID member channel
                generic.List<string> resultIdChannelGuids;

                //recupero tutti i membri di default
                var membersDefault = usersDefault.getAllMembers();
                var allMembers = membersDefault.Union(teamObject.membersChannel).ToList();
                resultIdChannelGuids = await Service.getIdUserFromEmail(graphClient, allMembers);


                string _NameTeamEmail = Regex.Replace(teamObject.NameTeamEmail, @" ", "");
                //creo il gruppo
                var resultGroup = await Service.createGroup(graphClient, teamObject.NameTeam, _NameTeamEmail);
                Thread.Sleep(5000);

                //aggiungo gli owners al gruppo
                await Service.addOwnerToGroup(graphClient, resultGroup.Id, resultGuidsOwner);
                Thread.Sleep(3000);

                //aggiungo i membri al gruppo
                await Service.addMemberToGroup(graphClient, resultGroup.Id, resultIdChannelGuids);
                Thread.Sleep(5000);
                //creo il team - quando creo il gruppo provato in automatico viene creato il team
                var resultTeam = await Service.createTeam(graphClient, resultGroup.Id, teamObject.NameTeam);

                var group = new GroupTeam();
                group.Id = resultGroup.Id;
                return group;

            }
            catch (System.Exception e)
            {
                //Console.WriteLine("Errore creazione del team " + e.Message);
                throw e;
            }
        }

        public static async Task createChannel(TeamCommessa teamObject)
        {
            try
            {
                GraphServiceClient graphClient = await Authentication.auth();
                //recupero i valori di default dei membri
                var token = Authentication.authSP().Result;
                usersDefault = JsonConvert.DeserializeObject<ConfigCanali>(Service.getValueConfigFile("ConfigUserChannels", token));

                var allOwners = removeDuplicateFromList(teamObject.ownersTeam);
                //recupero ID owner Team tramite email - owner team e canali soo gli stessi
                var resultGuidsOwner = await Service.getIdUserFromEmail(graphClient, allOwners);

                //creo i canali con all'interno almeno un owner
                foreach (var item in teamObject.Channels)
                {
                    try
                    {
                        var channel = await Service.createChannel(graphClient, teamObject.IdGroup, resultGuidsOwner[0], item.NameChannel);
                        item.IdChannel = channel.Id;
                        Console.WriteLine("--------Creazione canale " + item.NameChannel + " avvenuta con successo!");
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
                        await Service.addMoreOwnerToChannel(graphClient, teamObject.IdGroup, item.IdChannel, resultGuidsOwner);
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
                        var allDistinctmembers = await Service.getIdUserFromEmail(graphClient, allmembers);
                        await Service.addMemberToChannel(graphClient, teamObject.IdGroup, item.IdChannel, allDistinctmembers);
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

        public static void createObjectToListConfig(TeamCommessa objectTeam){
            try
            {
                var token = Authentication.authSP().Result;
                Service.insertObjectToListConfig(token,objectTeam);
            }
            catch (System.Exception)
            {
                throw new Exception("Errore Inserimento commessa nella lista di configurazione");
            }
        }

        public static void initializeStructureDL(TeamCommessa teamObject, string token)
        {
            var check = false;
            var UrlSite = "";
            try
            {
                //il canale general è l'unico che esiste sempre
                if(teamObject.StatoCreazione == "Strutture da creare"){
                    Service.updateObjectToListConfig(token,teamObject,"","In corso");
                    //recupero i valori di default 
                    jsonConfig = JsonConvert.DeserializeObject<ConfigFolder>(Service.getValueConfigFile("ConfigFile", token));
                    //imposto i content type,colonne e viste sul canale general
                    UrlSite = siteUrl + Regex.Replace(teamObject.NameTeam, @" ", "");
                    SettingsMajorUpdate(token, UrlSite, "Team");
                    CreateDocumentLibraryStructure(token, UrlSite, teamObject, "Team");
                    Console.WriteLine("Fine creazione struttura folder general");
                    Service.updateObjectToListConfig(token,teamObject,"Canale: General - OK","");
                }

                foreach (var channel in teamObject.Channels)
                {
                    try
                    {
                        if(channel.create != true){
                            UrlSite = siteUrl + Regex.Replace(teamObject.NameTeam, @" ", "") + "-" + Regex.Replace(channel.NameChannel, @" ", "");
                            SettingsMajorUpdate(token, UrlSite, "channel");
                            Console.WriteLine("Fine creazione struttura folder canale - " + channel.NameChannel);
                            CreateDocumentLibraryStructure(token, UrlSite, teamObject, channel.NameChannel);
                            Console.WriteLine("Fine impostazioni metadati canale- " + channel.NameChannel);
                            channel.create = true;
                            if(teamObject.StatoCreazione == "Da completare"){
                                Service.updateObjectToListConfig(token,teamObject,"***Canale:"+channel.NameChannel+"- OK",teamObject.StatoCreazione);
                            }else{
                                Service.updateObjectToListConfig(token,teamObject,"Canale:"+channel.NameChannel+"- OK",teamObject.StatoCreazione);
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
                            Service.updateObjectToListConfig(token,teamObject,"Canale:"+channel.NameChannel+"- struttura non creata","");
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
                    Service.updateObjectToListConfig(token,teamObject,"","Da completare");
                }else{
                    Service.updateObjectToListConfig(token,teamObject,"Strutture create con successo","Strutture create");
                }
            }
            catch (System.Exception e)
            {
                Service.updateObjectToListConfig(token,teamObject,e.Message,"Errore");
                throw e;
            }

        }

        public static void SettingsMajorUpdate(string token, string UrlSite,string checkTeam)
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
                    Documents.createContentType(context);
                    //creo le colonne nel site column e le aggiungo al content type
                    Documents.createColumns(context);
                    if(checkTeam == "Team")//aggiungo colonna commento pubblicazione
                    {
                        Documents.createTextOrChoiceColumn(context, "CommentiPubblicazione", "Commenti Pubblicazione", "multiline");
                    }
                    else
                    {
                        Documents.createTextOrChoiceColumn(context, "Stato", "Stato","choice");
                        Documents.createTextOrChoiceColumn(context,"Pubblica", "Pubblica", "flow");
                    }
                    //imposto il content type
                    Documents.setContentTypeToList(context);
                    //imposto le viste
                    Documents.setViewList(context,checkTeam);
                    Documents.createViewList(context, "Teams",checkTeam);
                }
                catch (Exception e)
                {
                    throw e;
                }

            }
        }

        public static void CreateDocumentLibraryStructure(string token, string UrlSite, TeamCommessa teamObject, string channel)
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
                            Documents.CreateFolder(context, UrlSite, relativePath, nameFolder);
                            //imposto i metadata
                            if (item.Name == "Years")
                            {
                                var idValueMetadato = Documents.getTaxonomyTermGroup(context, nameFolder);
                                if (idValueMetadato != null || idValueMetadato != "")
                                {
                                    Documents.setDefaulValueColumn(context, relativePath + "/" + nameFolder, item.Name + "MT", nameFolder + "|" + idValueMetadato);
                                }
                            }
                            else if (item.Name == "Customers")
                            {

                                var idValueMetadato = Documents.getTaxonomyTermGroup(context, nameFolder);
                                if (idValueMetadato != null || idValueMetadato != "")
                                {
                                    Documents.setDefaulValueColumn(context, relativePath + "/" + nameFolder, item.Name + "MT", nameFolder + "|" + idValueMetadato);
                                }

                                //imposto PM-canale
                                generic.List<string> allSettingCustomers = new generic.List<string>() { "PMs", "Companies", "ISO", "Cities" };
                                foreach (var val in teamObject.Metadata)
                                {
                                    if (allSettingCustomers.Contains(val.NameMetadato))
                                    {
                                        idValueMetadato = Documents.getTaxonomyTermGroup(context, val.ValueMetadato);
                                        if (idValueMetadato != null || idValueMetadato != "")
                                        {
                                            Documents.setDefaulValueColumn(context, relativePath + "/" + nameFolder, val.NameMetadato + "MT", val.ValueMetadato + "|" + idValueMetadato);
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

                                        var idValueMetadato = Documents.getTaxonomyTermGroup(context, val.ValueMetadato);
                                        if (idValueMetadato != null || idValueMetadato != "")
                                        {
                                            Documents.setDefaulValueColumn(context, relativePath + "/" + nameFolder, val.NameMetadato + "MT", val.ValueMetadato + "|" + idValueMetadato);
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
                                generic.List<string> labels = Documents.getValueTermSet(context, item.Name);
                                foreach (var label in labels)
                                {
                                    Documents.CreateFolder(context, UrlSite, relativePath, label);
                                    var idValueMetadato = Documents.getTaxonomyTermGroup(context, label);
                                    if (idValueMetadato != null || idValueMetadato != "")
                                    {
                                        Documents.setDefaulValueColumn(context, relativePath + "/" + label, "DocTypesMT", label + "|" + idValueMetadato);
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
