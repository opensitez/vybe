// vybe-test: csharp/csharp_span_indexing/span_from_array_slice_reads_correct_elements
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = { 10, 20, 30, 40, 50 };
var span = new System.Span<int>(data, 1, 3);
__Check((span[0]).ToString(), "20");
__Check((span[2]).ToString(), "40");
