// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_two_slices_compose_expected_element
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[6]{1,2,3,4,5,6}; var inner=span.Slice(2,2); __Check((inner[1]).ToString(), "4");
