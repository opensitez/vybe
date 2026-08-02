// vybe-test: csharp/csharp_decimal_financial/decimal_financial_parse_from_string
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((decimal.Parse("1234.56")).ToString(), "1234.56");
