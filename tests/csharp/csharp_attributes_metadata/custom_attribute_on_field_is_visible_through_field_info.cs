// vybe-test: csharp/csharp_attributes_metadata/custom_attribute_on_field_is_visible_through_field_info
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

using System; [AttributeUsage(AttributeTargets.Field)] class MarkerAttribute : Attribute { public int Code { get; } public MarkerAttribute(int code) { Code = code; } } class Flags { [Marker(7)] public int Value; } var field = typeof(Flags).GetField("Value"); var attr = (MarkerAttribute)Attribute.GetCustomAttribute(field, typeof(MarkerAttribute)); __P((attr.Code).ToString());
__Check("7");
