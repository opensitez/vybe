// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_memory_span_from_string_prints_first_char
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ReadOnlyMemory<char> mem="hello".AsMemory(); __Check((mem.Span[0]).ToString(), "104");
