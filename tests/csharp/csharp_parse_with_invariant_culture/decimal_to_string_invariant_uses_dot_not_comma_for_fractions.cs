// vybe-test: csharp/csharp_parse_with_invariant_culture/decimal_to_string_invariant_uses_dot_not_comma_for_fractions
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

decimal value = 2.25m;
__P((value.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToString());
__Check("2.25");
