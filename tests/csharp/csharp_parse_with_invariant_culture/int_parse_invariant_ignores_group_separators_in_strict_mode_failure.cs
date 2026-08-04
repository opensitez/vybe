// vybe-test: csharp/csharp_parse_with_invariant_culture/int_parse_invariant_ignores_group_separators_in_strict_mode_failure
// origin: languages/csharp/tests/csharp/test_csharp_parse_with_invariant_culture.rs

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

try {
    int.Parse("1,234", System.Globalization.CultureInfo.InvariantCulture);
    __P(("parsed").ToString());
} catch (System.FormatException) {
    __P(("reject").ToString());
}
__Check("reject");
