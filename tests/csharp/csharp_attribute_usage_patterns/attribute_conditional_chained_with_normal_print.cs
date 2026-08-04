// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_conditional_chained_with_normal_print
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

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

using System; using System.Diagnostics; class P{[Conditional("DEBUG")] static void D(){__P(("d").ToString());} static void N(){__P(("n").ToString());} public static void Go(){D(); N();}} P.Go();
__Check("n");
