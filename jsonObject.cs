using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Company.Function
{
      
    public class JsonObject    {
        [JsonProperty("$schema")]
        public string Schema { get; set; } 
        public string elmType { get; set; } 
        public CustomRowAction customRowAction { get; set; } 
        public Attributes attributes { get; set; } 
        public Style style { get; set; } 
        public List<Child> children { get; set; } 
    }
// Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 
    public class CustomRowAction    {
        public string action { get; set; } 
        public string actionParams { get; set; } 
    }

    public class Attributes    {
        public string @class { get; set; } 
    }

    public class Style    {
        public string border { get; set; } 
        [JsonProperty("background-color")]
        public string BackgroundColor { get; set; } 
        public string cursor { get; set; } 
        public string display { get; set; } 
    }

    public class Attributes2    {
        public string iconName { get; set; } 
    }

    public class Style2    {
        public string color { get; set; } 
        [JsonProperty("padding-right")]
        public string PaddingRight { get; set; } 
    }

    public class Child    {
        public string elmType { get; set; } 
        public Attributes2 attributes { get; set; } 
        public Style2 style { get; set; } 
        public string txtContent { get; set; } 
    }

  
  
}