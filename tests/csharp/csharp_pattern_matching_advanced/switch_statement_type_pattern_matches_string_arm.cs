// vybe-test: csharp/csharp_pattern_matching_advanced/switch_statement_type_pattern_matches_string_arm
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

object item = "beta"; switch (item) { case string text: __P((text.ToUpper()).ToString()); break; default: __P(("other").ToString()); break; }
__Check("BETA");
