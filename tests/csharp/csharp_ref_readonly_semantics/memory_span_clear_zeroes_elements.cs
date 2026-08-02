// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_clear_zeroes_elements
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory=new System.Memory<int>(new int[]{5,6}); memory.Span.Clear(); __Check((memory.Span[0]).ToString(), "0");
