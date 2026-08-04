// vybe-test: csharp/csharp_numeric_checked_bitwise/decimal_addition_preserves_decimal_precision
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

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

decimal left = 1.2m; decimal right = 2.3m; __P((left + right).ToString());
__Check("3.5");
