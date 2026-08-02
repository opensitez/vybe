// vybe-test: csharp/csharp_attributes_metadata/attribute_can_be_read_from_base_method_definition
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Method)] class InfoAttribute : Attribute { public string Name { get; } public InfoAttribute(string name) { Name = name; } } class Base { [Info("root")] public virtual void Run() { } } class Derived : Base { public override void Run() { } } var method = typeof(Base).GetMethod("Run"); var attr = (InfoAttribute)Attribute.GetCustomAttribute(method, typeof(InfoAttribute)); __Check((attr.Name).ToString(), "root");
