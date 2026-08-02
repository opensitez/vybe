// vybe-test: csharp/csharp_decimal_financial/decimal_financial_round_preserves_trailing_zero_scale
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((decimal.Round(3.10m,2).ToString("0.00")).ToString(), "3.10");
