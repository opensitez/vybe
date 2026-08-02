// vybe-test: csharp/csharp_parsing_formatting/string_format_replaces_indexed_placeholders
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("{0}-{1}", "A", 3)).ToString(), "A-3");
