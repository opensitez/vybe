// vybe-test: csharp/csharp_attributes/custom_attribute_readable_via_get_custom_attributes
// origin: languages/csharp/tests/csharp/test_csharp_attributes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.AttributeUsage(System.AttributeTargets.Class)]
class TagAttribute:System.Attribute{public string Value;public TagAttribute(string v){Value=v;}}
[Tag("hello")]
class Target{}
var attrs=(TagAttribute[])typeof(Target).GetCustomAttributes(typeof(TagAttribute),false);
__Check((attrs[0].Value).ToString(), "hello");
