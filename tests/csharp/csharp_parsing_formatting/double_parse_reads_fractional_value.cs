// vybe-test: csharp/csharp_parsing_formatting/double_parse_reads_fractional_value
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((double.Parse("3.5") + 0.5).ToString(), "4");
