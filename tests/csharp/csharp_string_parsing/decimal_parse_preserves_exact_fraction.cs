// vybe-test: csharp/csharp_string_parsing/decimal_parse_preserves_exact_fraction
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

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

var d=decimal.Parse("0.1",System.Globalization.CultureInfo.InvariantCulture);
__P((d+0.2m==0.3m).ToString());
__Check("True");
