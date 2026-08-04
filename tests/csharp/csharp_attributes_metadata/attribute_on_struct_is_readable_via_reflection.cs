// vybe-test: csharp/csharp_attributes_metadata/attribute_on_struct_is_readable_via_reflection
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

using System; [AttributeUsage(AttributeTargets.Struct)] class ShapeAttribute : Attribute { public string Name { get; } public ShapeAttribute(string name) { Name = name; } } [Shape("point")] struct Point { } var attr = (ShapeAttribute)Attribute.GetCustomAttribute(typeof(Point), typeof(ShapeAttribute)); __P((attr.Name).ToString());
__Check("point");
