// vybe-test: csharp/csharp_decimal_financial/decimal_financial_round_midpoint_to_even
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((decimal.Round(1.235m,2,System.MidpointRounding.ToEven)).ToString(), "1.24");
