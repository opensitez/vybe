// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_attribute_can_be_applied_by_short_name_after_using
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;
using Demo;

var attr = (Demo.TagAttribute)System.Attribute.GetCustomAttribute(typeof(Item), typeof(Demo.TagAttribute));
__P((attr.Name).ToString());
__Check("x");

namespace Demo { public class TagAttribute : System.Attribute { public string Name; public TagAttribute(string name) { Name = name; } } }

[Tag("x")] class Item { }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
