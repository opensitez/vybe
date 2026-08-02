// vybe-test: csharp/csharp_string_span/memory_span_property_accesses_underlying_data
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Memory<int> m=new int[]{7,8,9};
__Check((m.Span[1]).ToString(), "8");
