// vybe-test: csharp/csharp_namespace_aliases/namespace_can_contain_static_helper_class
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

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

namespace Demo.Tools { public static class MathEx { public static int Double(int value) { return value * 2; } } } __P((Demo.Tools.MathEx.Double(6)).ToString());
__Check("12");
