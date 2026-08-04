// vybe-test: csharp/csharp_decimal_financial/decimal_financial_divide_per_unit_cost
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

decimal bill=100.00m; decimal seats=6m; __P((bill/seats>16.6m&&bill/seats<16.7m).ToString());
__Check("True");
