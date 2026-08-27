// vybe-test: csharp/csharp_attributes/custom_attribute_readable_via_get_custom_attributes
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

using static __Harness;

var attrs=(TagAttribute[])typeof(Target).GetCustomAttributes(typeof(TagAttribute),false);
__P((attrs[0].Value).ToString());
__Check("hello");

[System.AttributeUsage(System.AttributeTargets.Class)]
class TagAttribute:System.Attribute{public string Value;public TagAttribute(string v){Value=v;}}

[Tag("hello")]
class Target{}

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
