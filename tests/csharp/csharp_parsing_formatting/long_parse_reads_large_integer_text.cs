// vybe-test: csharp/csharp_parsing_formatting/long_parse_reads_large_integer_text
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((long.Parse("123456") + 1).ToString(), "123457");
