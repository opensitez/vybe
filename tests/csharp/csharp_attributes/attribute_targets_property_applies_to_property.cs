// vybe-test: csharp/csharp_attributes/attribute_targets_property_applies_to_property
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

using static __Harness;

var pi=typeof(Form).GetProperty("Name");
bool has=pi.GetCustomAttributes(typeof(RequiredAttribute),false).Length>0;
__P((has).ToString());
__Check("True");

[System.AttributeUsage(System.AttributeTargets.Property)]
class RequiredAttribute:System.Attribute{}

class Form{[Required] public string Name{get;set;}}

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
