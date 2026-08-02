// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_from_array_as_span_then_slice
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3,4}; System.Span<int> span=data.AsSpan(1,2); __Check((span[0]).ToString(), "2"); __Check((span[1]).ToString(), "3");
