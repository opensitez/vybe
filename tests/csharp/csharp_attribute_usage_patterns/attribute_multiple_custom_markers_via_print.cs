// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_multiple_custom_markers_via_print
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Class)] class TagAttribute:Attribute{public string Name; public TagAttribute(string n){Name=n;}} [Tag("svc")] class Worker{} var a=(TagAttribute)Attribute.GetCustomAttribute(typeof(Worker),typeof(TagAttribute)); __Check((a.Name).ToString(), "svc");
