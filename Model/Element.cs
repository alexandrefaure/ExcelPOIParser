using Newtonsoft.Json;

namespace TestExcelParser.Model
{
    public class Element
    {
        [JsonProperty("Cell")] public Cell cell;

        public string structuredValue;
        public string verificationValue;
    }
}