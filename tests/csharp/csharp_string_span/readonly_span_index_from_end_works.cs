// vybe-test: csharp/csharp_string_span/readonly_span_index_from_end_works
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ReadOnlySpan<char> s="hello".AsSpan();
__Check((s[^1]).ToString(), "o");
