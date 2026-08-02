// vybe-test: csharp/csharp_parsing_formatting/decimal_parse_reads_decimal_literal_text
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((decimal.Parse("7.25")).ToString(), "7.25");
