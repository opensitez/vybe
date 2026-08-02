// vybe-test: csharp/csharp_parsing_formatting/float_parse_reads_fractional_text
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((float.Parse("2.5") + 0.5f).ToString(), "3");
