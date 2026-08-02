// vybe-test: csharp/csharp_decimal_financial/decimal_financial_weighted_average
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal w1=0.6m; decimal w2=0.4m; decimal p1=10m; decimal p2=20m; __Check((w1*p1+w2*p2).ToString(), "14.0");
