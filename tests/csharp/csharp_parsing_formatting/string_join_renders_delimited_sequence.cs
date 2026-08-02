// vybe-test: csharp/csharp_parsing_formatting/string_join_renders_delimited_sequence
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Join("|", new[] { "a", "b", "c" })).ToString(), "a|b|c");
