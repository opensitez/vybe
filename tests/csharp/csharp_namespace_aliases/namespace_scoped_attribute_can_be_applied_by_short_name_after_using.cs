// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_attribute_can_be_applied_by_short_name_after_using
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; namespace Demo { public class TagAttribute : System.Attribute { public string Name; public TagAttribute(string name) { Name = name; } } } [Tag("x")] class Item { } var attr = (Demo.TagAttribute)System.Attribute.GetCustomAttribute(typeof(Item), typeof(Demo.TagAttribute)); __P((attr.Name).ToString());
__Check("x");
