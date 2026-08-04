// vybe-test: csharp/csharp_attributes_metadata/attribute_constructor_can_capture_integer_argument
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

using System; [AttributeUsage(AttributeTargets.Class)] class CodeAttribute : Attribute { public int Value { get; } public CodeAttribute(int value) { Value = value; } } [Code(42)] class Job { } var attr = (CodeAttribute)Attribute.GetCustomAttribute(typeof(Job), typeof(CodeAttribute)); __P((attr.Value).ToString());
__Check("42");
