// vybe-test: csharp/csharp_decimal_financial/decimal_financial_equality_ignores_trailing_zeros
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((2.50m==2.5m).ToString(), "True");
