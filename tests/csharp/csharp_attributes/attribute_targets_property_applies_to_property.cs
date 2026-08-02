// vybe-test: csharp/csharp_attributes/attribute_targets_property_applies_to_property
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.AttributeUsage(System.AttributeTargets.Property)]
class RequiredAttribute:System.Attribute{}
class Form{[Required] public string Name{get;set;}}
var pi=typeof(Form).GetProperty("Name");
bool has=pi.GetCustomAttributes(typeof(RequiredAttribute),false).Length>0;
__Check((has).ToString(), "True");
