// vybe-test: csharp/csharp_decimal_financial/decimal_financial_amortized_payment_estimate
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal loan=12000m; decimal months=12m; __Check((loan/months).ToString(), "1000");
