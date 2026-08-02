// vybe-test: csharp/csharp_index_from_end/range_to_end_from_index_from_end_produces_tail_slice
// origin: languages/csharp/tests/csharp/test_csharp_index_from_end.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = { 1, 2, 3, 4 };
var tail = data[2..^0];
__Check((tail.Length).ToString(), "2");
__Check((tail[0]).ToString(), "3");
