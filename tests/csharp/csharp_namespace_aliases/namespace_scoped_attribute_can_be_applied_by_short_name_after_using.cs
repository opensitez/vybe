// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_attribute_can_be_applied_by_short_name_after_using
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public class TagAttribute : System.Attribute { public string Name; public TagAttribute(string name) { Name = name; } } } [Tag("x")] class Item { } var attr = (Demo.TagAttribute)System.Attribute.GetCustomAttribute(typeof(Item), typeof(Demo.TagAttribute)); __Check((attr.Name).ToString(), "x");
