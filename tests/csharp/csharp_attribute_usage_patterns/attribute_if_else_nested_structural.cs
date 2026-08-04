// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_if_else_nested_structural
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

#define VYBETEST_A
#if VYBETEST_A
__P(("a").ToString());
#else
__P(("b").ToString());
#endif
__P(("c").ToString());
__Check("a\nc");
