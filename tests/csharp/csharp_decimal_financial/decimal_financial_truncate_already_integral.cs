// vybe-test: csharp/csharp_decimal_financial/decimal_financial_truncate_already_integral
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((decimal.Truncate(42.00m)).ToString(), "42");
