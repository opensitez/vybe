// vybe-test: csharp/csharp_attributes_metadata/custom_attribute_on_field_is_visible_through_field_info
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Field)] class MarkerAttribute : Attribute { public int Code { get; } public MarkerAttribute(int code) { Code = code; } } class Flags { [Marker(7)] public int Value; } var field = typeof(Flags).GetField("Value"); var attr = (MarkerAttribute)Attribute.GetCustomAttribute(field, typeof(MarkerAttribute)); __Check((attr.Code).ToString(), "7");
