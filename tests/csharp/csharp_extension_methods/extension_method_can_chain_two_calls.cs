// vybe-test: csharp/csharp_extension_methods/extension_method_can_chain_two_calls
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

using Demo; namespace Demo { public static class TextExt { public static string AddA(this string value) { return value + "a"; } public static string AddB(this string value) { return value + "b"; } } } __P(("x".AddA().AddB()).ToString());
__Check("xab");
