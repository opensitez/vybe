// vybe-test: csharp/csharp_attributes_metadata/attribute_allow_multiple_returns_two_instances
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

using System; [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } } [Tag("api"), Tag("internal")] class Endpoint { } var attrs = typeof(Endpoint).GetCustomAttributes(typeof(TagAttribute), false); __P((attrs.Length).ToString());
__Check("2");
