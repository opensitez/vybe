// vybe-test: csharp/csharp_attributes_metadata/custom_attribute_on_property_is_visible_through_property_info
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Property)] class HintAttribute : Attribute { public string Text { get; } public HintAttribute(string text) { Text = text; } } class Settings { [Hint("port")] public int Port { get; set; } } var property = typeof(Settings).GetProperty("Port"); var attr = (HintAttribute)Attribute.GetCustomAttribute(property, typeof(HintAttribute)); __Check((attr.Text).ToString(), "port");
