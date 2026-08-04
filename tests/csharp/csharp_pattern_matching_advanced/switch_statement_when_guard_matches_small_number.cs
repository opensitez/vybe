// vybe-test: csharp/csharp_pattern_matching_advanced/switch_statement_when_guard_matches_small_number
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

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

var x = 3; switch (x) { case int number when number > 10: __P(("large").ToString()); break; case int number: __P(("small").ToString()); break; }
__Check("small");
