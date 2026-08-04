// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_multiple_custom_markers_via_print
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

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

using System; [AttributeUsage(AttributeTargets.Class)] class TagAttribute:Attribute{public string Name; public TagAttribute(string n){Name=n;}} [Tag("svc")] class Worker{} var a=(TagAttribute)Attribute.GetCustomAttribute(typeof(Worker),typeof(TagAttribute)); __P((a.Name).ToString());
__Check("svc");
