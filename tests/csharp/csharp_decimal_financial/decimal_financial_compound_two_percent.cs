// vybe-test: csharp/csharp_decimal_financial/decimal_financial_compound_two_percent
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

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

decimal principal=1000.00m; decimal rate=0.02m; __P((principal*(1m+rate)).ToString());
__Check("1020.00");
