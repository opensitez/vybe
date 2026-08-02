// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_slice_preserves_element_after_mutation
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[4]{1,2,3,4}; var part=span.Slice(1,2); part[0]=88; __Check((span[1]).ToString(), "88");
