// vybe-test: csharp/csharp_attributes_metadata/custom_attribute_on_class_can_be_read_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Class)] class LabelAttribute : Attribute { public string Name { get; } public LabelAttribute(string name) { Name = name; } } [Label("service")] class Worker { } var attr = (LabelAttribute)Attribute.GetCustomAttribute(typeof(Worker), typeof(LabelAttribute)); __Check((attr.Name).ToString(), "service");
