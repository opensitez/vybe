// vybe-test: csharp/csharp_attributes/multiple_attributes_on_same_target_both_retrieved
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.AttributeUsage(System.AttributeTargets.Class,AllowMultiple=true)]
class TagAttribute:System.Attribute{public string Name;public TagAttribute(string n){Name=n;}}
[Tag("a")][Tag("b")]
class Thing{}
var attrs=(TagAttribute[])typeof(Thing).GetCustomAttributes(typeof(TagAttribute),false);
__Check((attrs.Length).ToString(), "2");
