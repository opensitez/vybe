// vybe-test: csharp/csharp_extension_methods/generic_extension_method_returns_same_value
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

using Demo; namespace Demo { public static class EchoExt { public static T Echo<T>(this T value) { return value; } } } __P(("hi".Echo()).ToString());
__Check("hi");
