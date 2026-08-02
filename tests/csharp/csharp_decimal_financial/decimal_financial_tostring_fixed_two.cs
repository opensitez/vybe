// vybe-test: csharp/csharp_decimal_financial/decimal_financial_tostring_fixed_two
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((7.5m.ToString("F2")).ToString(), "7.50");
