// vybe-test: csharp/csharp_attributes_metadata/attribute_get_custom_attributes_can_return_strongly_typed_array
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

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

using System; [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } } [Tag("a"), Tag("b")] class Demo { } var attrs = (TagAttribute[])typeof(Demo).GetCustomAttributes(typeof(TagAttribute), false); foreach (var attr in attrs) __P((attr.Name).ToString());
__Check("a\nb");
