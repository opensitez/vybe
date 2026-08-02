// vybe-test: csharp/csharp_decimal_financial/decimal_financial_margin_calculation
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal revenue=500m; decimal cost=320m; __Check(((revenue-cost)/revenue).ToString(), "0.36");
