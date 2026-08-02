// vybe-test: csharp/csharp_stackalloc_span/stackalloc_float_buffer_reads_single_precision_value
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<float> span=stackalloc float[2]{1.25f,2.5f}; __Check((span[0]==1.25f).ToString(), "True");
