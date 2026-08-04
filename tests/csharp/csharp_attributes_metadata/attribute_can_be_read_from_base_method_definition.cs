// vybe-test: csharp/csharp_attributes_metadata/attribute_can_be_read_from_base_method_definition
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

using System; [AttributeUsage(AttributeTargets.Method)] class InfoAttribute : Attribute { public string Name { get; } public InfoAttribute(string name) { Name = name; } } class Base { [Info("root")] public virtual void Run() { } } class Derived : Base { public override void Run() { } } var method = typeof(Base).GetMethod("Run"); var attr = (InfoAttribute)Attribute.GetCustomAttribute(method, typeof(InfoAttribute)); __P((attr.Name).ToString());
__Check("root");
