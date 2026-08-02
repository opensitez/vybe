// vybe-test: csharp/csharp_decimal_financial/decimal_financial_compareto_greater
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((5.0m.CompareTo(4.9m)).ToString(), "1");
