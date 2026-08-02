// vybe-test: csharp/csharp_decimal_financial/decimal_financial_compareto_equal_scale
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((1.0m.CompareTo(1.00m)).ToString(), "0");
