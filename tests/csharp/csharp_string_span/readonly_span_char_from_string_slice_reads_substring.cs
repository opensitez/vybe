// vybe-test: csharp/csharp_string_span/readonly_span_char_from_string_slice_reads_substring
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s="hello world";
System.ReadOnlySpan<char> span=s.AsSpan(6,5);
__Check((span.ToString()).ToString(), "world");
