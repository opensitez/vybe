// vybe-test: csharp/csharp_numeric_formatting/format_custom_zero_placeholder_pads_fractional_digits
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

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

__P(((1.5).ToString("0.00")).ToString());
__Check("1.50");
