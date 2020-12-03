using System.Collections.Generic;

namespace Company.Function
{
    class ConfigFolder
    {
        public List<Folder> LevelDescriptors { get; set; }
    }
    class Folder
    {
        //tipo di dato - esempio Sub Suddivisione è un tipo di dato Testo e non termine
        public ValueTypes Type{ get; set; }
        //nome del metadato
        public string Name { get; set; }
        public string Level { get; set; }
        public bool Leaves{get;set;}
        //nel caso di Customer (metadato) abbiamo un nesting del metadato pari a Allianz[posizizione 1]-Allianz[posizione 2]
        public string PositionTerm { get; set; }
    }


    enum ValueTypes
    {
        Term,
        Text
    }
}
