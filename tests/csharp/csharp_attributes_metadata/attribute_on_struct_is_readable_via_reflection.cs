// vybe-test: csharp/csharp_attributes_metadata/attribute_on_struct_is_readable_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Struct)] class ShapeAttribute : Attribute { public string Name { get; } public ShapeAttribute(string name) { Name = name; } } [Shape("point")] struct Point { } var attr = (ShapeAttribute)Attribute.GetCustomAttribute(typeof(Point), typeof(ShapeAttribute)); __Check((attr.Name).ToString(), "point");
