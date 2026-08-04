// vybe-test: csharp/csharp_extension_methods/extension_method_on_nullable_int_handles_has_value
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

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

using Demo; namespace Demo { public static class NullableExt { public static string Describe(this int? value) { return value.HasValue ? value.Value.ToString() : "none"; } } } int? value = 8; __P((value.Describe()).ToString());
__Check("8");
