// vybe-test: csharp/csharp_attributes_metadata/attribute_inheritance_flows_to_derived_type_when_enabled
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

using System; [AttributeUsage(AttributeTargets.Class, Inherited = true)] class RoleAttribute : Attribute { public string Name { get; } public RoleAttribute(string name) { Name = name; } } [Role("base")] class BaseController { } class DerivedController : BaseController { } var attr = (RoleAttribute)Attribute.GetCustomAttribute(typeof(DerivedController), typeof(RoleAttribute)); __P((attr.Name).ToString());
__Check("base");
