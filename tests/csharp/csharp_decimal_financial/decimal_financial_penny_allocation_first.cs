// vybe-test: csharp/csharp_decimal_financial/decimal_financial_penny_allocation_first
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

decimal total=0.10m; int parts=3; decimal share=decimal.Truncate(total/parts*100m)/100m; __P((share).ToString());
__Check("0.03");
