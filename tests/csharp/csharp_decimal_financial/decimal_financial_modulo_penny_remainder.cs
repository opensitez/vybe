// vybe-test: csharp/csharp_decimal_financial/decimal_financial_modulo_penny_remainder
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((10.01m%0.10m).ToString(), "0.01");
