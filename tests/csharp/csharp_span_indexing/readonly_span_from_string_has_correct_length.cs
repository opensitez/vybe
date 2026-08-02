// vybe-test: csharp/csharp_span_indexing/readonly_span_from_string_has_correct_length
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ReadOnlySpan<char> span = "hello".AsSpan();
__Check((span.Length).ToString(), "5");
