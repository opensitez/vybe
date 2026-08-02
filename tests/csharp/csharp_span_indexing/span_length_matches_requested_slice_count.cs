// vybe-test: csharp/csharp_span_indexing/span_length_matches_requested_slice_count
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = { 1, 2, 3, 4 };
var span = data.AsSpan(1, 2);
__Check((span.Length).ToString(), "2");
