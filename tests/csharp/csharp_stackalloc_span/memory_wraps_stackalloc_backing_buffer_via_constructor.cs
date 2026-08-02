// vybe-test: csharp/csharp_stackalloc_span/memory_wraps_stackalloc_backing_buffer_via_constructor
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Memory<int> mem=new System.Memory<int>(stackalloc int[2]{1,2}); __Check((mem.Length).ToString(), "2");
