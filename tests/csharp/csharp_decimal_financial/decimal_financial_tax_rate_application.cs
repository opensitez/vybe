// vybe-test: csharp/csharp_decimal_financial/decimal_financial_tax_rate_application
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal price=100.00m; decimal rate=0.0825m; __Check((price*rate).ToString(), "8.2500");
