using System.Collections.Generic;

namespace Company.Function
{
    class ConfigCanali{
        public List<Canale> Canali { get; set; }

        public List<string> getAllMembers() {
            List<string> allMembers = new List<string>();
            foreach (var item in this.Canali)
            {
                foreach (var member in item.Members)
                {
                    allMembers.Add(member);
                }
            }
            return allMembers;
        }

        public List<string> getMembersChannel(string nameChannel) {
            List<string> allMembers = new List<string>();
            foreach (var item in this.Canali)
            {
                if(item.NameChannel == nameChannel){
                    foreach (var member in item.Members)
                    {
                        allMembers.Add(member);
                    }
                    break;
                }
                
            }
            return allMembers;
        }

    }
    class Canale
    {
        //Canali: Ufficio Tecnico, Customer Service - PM,Ufficio Tecnico, Order Entry, 
        //Produzione Acquisti,Logistica - Installazione - Montaggio

        //nome canale all'interno del team
        public string NameChannel { get; set; }
        //Members channel
        public List<string> Members{ get; set; }
    }
       
}