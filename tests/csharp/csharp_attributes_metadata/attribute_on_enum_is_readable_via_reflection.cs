// vybe-test: csharp/csharp_attributes_metadata/attribute_on_enum_is_readable_via_reflection
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

using System; [AttributeUsage(AttributeTargets.Enum)] class GroupAttribute : Attribute { public string Name { get; } public GroupAttribute(string name) { Name = name; } } [Group("status")] enum State { Idle } var attr = (GroupAttribute)Attribute.GetCustomAttribute(typeof(State), typeof(GroupAttribute)); __P((attr.Name).ToString());
__Check("status");
