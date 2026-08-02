// vybe-test: csharp/csharp_attributes_metadata/attribute_on_nested_class_is_readable_via_type_handle
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Class)] class LabelAttribute : Attribute { public string Name { get; } public LabelAttribute(string name) { Name = name; } } class Outer { [Label("inner")] public class Inner { } } var attr = (LabelAttribute)Attribute.GetCustomAttribute(typeof(Outer.Inner), typeof(LabelAttribute)); __Check((attr.Name).ToString(), "inner");
