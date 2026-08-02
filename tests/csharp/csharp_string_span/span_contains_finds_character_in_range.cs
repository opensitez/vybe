// vybe-test: csharp/csharp_string_span/span_contains_finds_character_in_range
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ReadOnlySpan<char> span="hello".AsSpan();
__Check((span.Contains('e')).ToString(), "True");
