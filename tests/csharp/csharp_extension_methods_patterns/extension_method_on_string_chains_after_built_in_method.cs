// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_on_string_chains_after_built_in_method
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

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

static class StrExt { public static string Shout(this string s) => s.ToUpper() + "!"; }
__P(("hello".Shout()).ToString());
__Check("HELLO!");
