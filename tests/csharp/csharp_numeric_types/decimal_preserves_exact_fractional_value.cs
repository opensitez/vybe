// vybe-test: csharp/csharp_numeric_types/decimal_preserves_exact_fractional_value
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

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

decimal d = 0.1m + 0.2m; __P((d).ToString());
__Check("0.3");
