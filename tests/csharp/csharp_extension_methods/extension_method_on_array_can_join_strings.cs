// vybe-test: csharp/csharp_extension_methods/extension_method_on_array_can_join_strings
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

using Demo; namespace Demo { public static class StringArrayExt { public static string JoinAll(this string[] values) { return string.Join(",", values); } } } __P((new[] { "a", "b" }.JoinAll()).ToString());
__Check("a,b");
