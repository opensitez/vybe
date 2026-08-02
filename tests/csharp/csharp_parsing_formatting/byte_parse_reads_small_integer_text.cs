// vybe-test: csharp/csharp_parsing_formatting/byte_parse_reads_small_integer_text
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((byte.Parse("12") + 1).ToString(), "13");
