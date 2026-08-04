// vybe-test: csharp/csharp_attributes/multiple_attributes_on_same_target_both_retrieved
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

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

[System.AttributeUsage(System.AttributeTargets.Class,AllowMultiple=true)]
class TagAttribute:System.Attribute{public string Name;public TagAttribute(string n){Name=n;}}
[Tag("a")][Tag("b")]
class Thing{}
var attrs=(TagAttribute[])typeof(Thing).GetCustomAttributes(typeof(TagAttribute),false);
__P((attrs.Length).ToString());
__Check("2");
