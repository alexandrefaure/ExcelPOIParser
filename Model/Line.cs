using System.Collections.Generic;
using Newtonsoft.Json;

namespace TestExcelParser.Model
{
    public class Line
    {
        [JsonProperty("ElementGroupsList")] public List<ElementGroup> ElementGroupsList;

        public string Level;
        public string Position;
        public string Type;
        public string UniqueId;
    }
}