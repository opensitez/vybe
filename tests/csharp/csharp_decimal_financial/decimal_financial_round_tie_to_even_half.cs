// vybe-test: csharp/csharp_decimal_financial/decimal_financial_round_tie_to_even_half
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((decimal.Round(2.5m,0)).ToString(), "2");
