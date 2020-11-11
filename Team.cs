using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Company.Function
{
  class Team
  {
    //nome del team corrisponde anche al nome del gruppo
    public string NameTeam { get; set; }
    //Canale : Customer Service - PM
    public string NameChannel1 { get; set; }
    //Canale : Ufficio Tecnico
    public string NameChannel2 { get; set; }
    //Canale : Order Entry
    public string NameChannel3 { get; set; }
    //Canale : Produzione Acquisti
    public string NameChannel4 { get; set; }
    //Canale : Logistica - Installazione - Montaggio
    public string NameChannel5 { get; set; }
    //Owner del team
    public string ownerTeam { get; set; }
    //Owner canali
    public string ownerChannel { get; set; }
    public List<string> listMemberChannel1 { get; set; }
    public List<string> listMemberChannel2 { get; set; }
    public List<string> listMemberChannel3 { get; set; }
    public List<string> listMemberChannel4 { get; set; }
    public List<string> listMemberChannel5 { get; set; }
    //METADATI
    //canale = companies
    public string Canale { get; set; }
    //project manager = PMs
    public string PMs { get; set; }
    //Main project name = Customers (Livello 0:es-Allianz)
    public string MainProjectName { get; set; }
    //Sub project name = Customers (Livello 1:es-Genialloyd)
    public string SubProjectName { get; set; }
    //Sub suddivisione = Customers (Livello 2:es-Piano1)
    public string SubSuddivisione { get; set; }
    //Anno = Years
    public string Years { get; set; }
    //ISO = ISO
    public string ISO { get; set; }
    //Citta = Cities
    public string Cities { get; set; }
    //Numero Offerta = numero offerta
    public string NumberOffer { get; set; }
  }
}
