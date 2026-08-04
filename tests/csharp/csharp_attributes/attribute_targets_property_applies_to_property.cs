// vybe-test: csharp/csharp_attributes/attribute_targets_property_applies_to_property
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

[System.AttributeUsage(System.AttributeTargets.Property)]
class RequiredAttribute:System.Attribute{}
class Form{[Required] public string Name{get;set;}}
var pi=typeof(Form).GetProperty("Name");
bool has=pi.GetCustomAttributes(typeof(RequiredAttribute),false).Length>0;
__P((has).ToString());
__Check("True");
