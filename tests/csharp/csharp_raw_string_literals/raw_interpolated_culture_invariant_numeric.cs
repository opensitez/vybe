// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_culture_invariant_numeric
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

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

double value=1234.5; string text=$"""{value.ToString(System.Globalization.CultureInfo.InvariantCulture)}"""; __P((text.Contains(".")).ToString());
__Check("True");
