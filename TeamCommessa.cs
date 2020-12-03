using System;
using System.Collections.Generic;

namespace Company.Function
{
  public class TeamCommessa
  {
    //nome del team corrisponde anche al nome del gruppo
    //riguardo al progetto sarebbe - nome del progetto -  nella form
    public string IdGroup { get; set; }
    public string NameTeam { get; set; }
    public string NameTeamEmail { get; set; }
    public string StatoCreazione { get; set; }
    public List<string> ownersTeam { get; set; }
    public List<string> membersChannel { get; set; }
    //canali
    public List<ChannelTeam> Channels { get; set; }
    public List<Metadati> Metadata { get; set; }
    
  }

  public class ChannelTeam{
    //Canali: Ufficio Tecnico, Customer Service - PM,Ufficio Tecnico, Order Entry, 
    //Produzione Acquisti,Logistica - Installazione - Montaggio

    //nome canale all'interno del team
    public string NameChannel { get; set; }
    public string IdChannel { get;set;}
    //Members channel
    public List<string> Members{ get; set; }
    public bool create{get;set;}
  }
  
  public class Metadati{
    //canale = companies
    //project manager = PMs
    //Main project name = Customers (Livello 0:es-Allianz + Livello1:es-Genialloyd)
    //Sub suddivisione = Customers (Livello 2:es-Piano1)
    //Anno = Years
    //ISO = ISO
    //Citta = Cities
    //Numero Offerta = numero offerta
    public string NameMetadato { get; set; }
    public string ValueMetadato { get; set; }
  }

  public class GroupTeam{
    public string Id { get; set; }


  }

}
