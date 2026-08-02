// vybe-test: csharp/csharp_attributes_metadata/custom_attribute_on_method_can_be_read_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Method)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } } class Worker { [Tag("run")] public void Execute() { } } var method = typeof(Worker).GetMethod("Execute"); var attr = (TagAttribute)Attribute.GetCustomAttribute(method, typeof(TagAttribute)); __Check((attr.Name).ToString(), "run");
