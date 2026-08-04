// vybe-test: csharp/csharp_decimal_semantics/decimal_subtraction_yields_exact_difference_for_currency_style_values
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

decimal price = 19.99m; decimal discount = 4.50m; __P((price - discount).ToString());
__Check("15.49");
