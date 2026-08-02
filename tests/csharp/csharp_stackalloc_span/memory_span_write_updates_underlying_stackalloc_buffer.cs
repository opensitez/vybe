// vybe-test: csharp/csharp_stackalloc_span/memory_span_write_updates_underlying_stackalloc_buffer
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Memory<int> mem=new System.Memory<int>(stackalloc int[2]{1,2}); mem.Span[0]=77; __Check((mem.Span[0]).ToString(), "77");
