// vybe-test: csharp/csharp_decimal_financial/decimal_financial_discount_percentage
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal list=250.00m; decimal pct=0.20m; __Check((list*(1m-pct)).ToString(), "200.00");
