// vybe-test: csharp/csharp_stackalloc_span/stackalloc_bool_buffer_stores_true_literal
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<bool> span=stackalloc bool[2]{true,false}; __Check((span[0]).ToString(), "True");
