using System.Collections.Generic;
using Newtonsoft.Json;

namespace TestExcelParser.Model
{
    public class ElementGroup
    {
        [JsonProperty("ElementsList")] public List<Element> elementsList;

        public string groupName;
        public string isOutOfContext;
    }
}