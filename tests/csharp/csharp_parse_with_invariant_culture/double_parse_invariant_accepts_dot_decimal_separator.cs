// vybe-test: csharp/csharp_parse_with_invariant_culture/double_parse_invariant_accepts_dot_decimal_separator
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

double value = double.Parse("3.5", System.Globalization.CultureInfo.InvariantCulture);
__P((value).ToString());
__Check("3.5");
