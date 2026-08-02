// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_fill_sets_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]; span.Fill(9); __Check((span[0]).ToString(), "9"); __Check((span[2]).ToString(), "9");
