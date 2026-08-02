// vybe-test: csharp/csharp_decimal_financial/decimal_financial_round_three_decimal_places
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((decimal.Round(0.1235m,3)).ToString(), "0.124");
