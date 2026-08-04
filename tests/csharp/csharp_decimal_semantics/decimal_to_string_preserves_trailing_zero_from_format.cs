// vybe-test: csharp/csharp_decimal_semantics/decimal_to_string_preserves_trailing_zero_from_format
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

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

decimal value = 3.5m; __P((value.ToString("0.00")).ToString());
__Check("3.50");
