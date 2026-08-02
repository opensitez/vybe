// vybe-test: csharp/csharp_span_indexing/span_write_mutates_backing_array
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = { 1, 2, 3 };
var span = data.AsSpan();
span[1] = 99;
__Check((data[1]).ToString(), "99");
