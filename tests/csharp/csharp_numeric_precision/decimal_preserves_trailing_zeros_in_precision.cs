// vybe-test: csharp/csharp_numeric_precision/decimal_preserves_trailing_zeros_in_precision
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

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

decimal d=1.50m;
__P((d.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToString());
__Check("1.50");
