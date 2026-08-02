// vybe-test: csharp/csharp_decimal_semantics/decimal_parse_reads_literal_text_without_exponent
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal value = decimal.Parse("42.5"); __Check((value).ToString(), "42.5");
