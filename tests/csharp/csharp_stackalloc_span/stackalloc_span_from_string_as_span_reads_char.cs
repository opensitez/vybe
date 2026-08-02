// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_from_string_as_span_reads_char
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ReadOnlySpan<char> chars="abcd".AsSpan(1,2); __Check((chars[0]).ToString(), "b"); __Check((chars[1]).ToString(), "c");
