// vybe-test: csharp/csharp_stackalloc_span/memory_span_reads_element_from_stackalloc_backing
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Memory<int> mem=new System.Memory<int>(stackalloc int[3]{4,5,6}); __Check((mem.Span[1]).ToString(), "5");
