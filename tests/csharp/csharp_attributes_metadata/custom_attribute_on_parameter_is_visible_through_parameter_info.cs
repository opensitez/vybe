// vybe-test: csharp/csharp_attributes_metadata/custom_attribute_on_parameter_is_visible_through_parameter_info
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

using System; [AttributeUsage(AttributeTargets.Parameter)] class UnitAttribute : Attribute { public string Name { get; } public UnitAttribute(string name) { Name = name; } } class MathOps { public void Scale([Unit("px")] int value) { } } var parameter = typeof(MathOps).GetMethod("Scale").GetParameters()[0]; var attr = (UnitAttribute)Attribute.GetCustomAttribute(parameter, typeof(UnitAttribute)); __P((attr.Name).ToString());
__Check("px");
