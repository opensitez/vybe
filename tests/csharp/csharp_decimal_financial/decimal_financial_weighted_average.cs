// vybe-test: csharp/csharp_decimal_financial/decimal_financial_weighted_average
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

decimal w1=0.6m; decimal w2=0.4m; decimal p1=10m; decimal p2=20m; __P((w1*p1+w2*p2).ToString());
__Check("14.0");
