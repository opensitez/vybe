// vybe-test: csharp/csharp_decimal_financial/decimal_financial_compound_two_percent
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal principal=1000.00m; decimal rate=0.02m; __Check((principal*(1m+rate)).ToString(), "1020.00");
