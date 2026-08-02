// vybe-test: csharp/csharp_span_indexing/memory_slice_reads_correct_element_via_span
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory = new System.Memory<int>(new int[] { 5, 6, 7 });
__Check((memory.Span[2]).ToString(), "7");
