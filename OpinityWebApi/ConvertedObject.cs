using Newtonsoft.Json;
using System.Dynamic;

namespace OpinityWebApi
{
    [JsonObject]
    public class ConvertedObject : DynamicObject
    {
        [JsonExtensionData]
        public Dictionary<string, object> DynamicProperties { get; set; }

        public ConvertedObject() 
        {
            DynamicProperties = new Dictionary<string, object>();
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            DynamicProperties.Add(binder.Name, value);

            return true;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            return DynamicProperties.TryGetValue(binder.Name, out result);
        }

        public void AddProperty(string name, object value)
        {
            DynamicProperties[name] = value;
        }
    }
}
