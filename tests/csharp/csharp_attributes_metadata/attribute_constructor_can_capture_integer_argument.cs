// vybe-test: csharp/csharp_attributes_metadata/attribute_constructor_can_capture_integer_argument
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Class)] class CodeAttribute : Attribute { public int Value { get; } public CodeAttribute(int value) { Value = value; } } [Code(42)] class Job { } var attr = (CodeAttribute)Attribute.GetCustomAttribute(typeof(Job), typeof(CodeAttribute)); __Check((attr.Value).ToString(), "42");
