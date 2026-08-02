// vybe-test: csharp/csharp_decimal_financial/decimal_financial_max_of_two_quotes
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Max(12.34m,12.35m)).ToString(), "12.35");
