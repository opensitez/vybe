// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_if_defined_symbol_prints_on_branch
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

#define VYBETEST_ON
#if VYBETEST_ON
__P(("on").ToString());
#else
__P(("off").ToString());
#endif
__Check("on");
