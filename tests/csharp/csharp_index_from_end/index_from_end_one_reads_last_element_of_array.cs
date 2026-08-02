// vybe-test: csharp/csharp_index_from_end/index_from_end_one_reads_last_element_of_array
// origin: languages/csharp/tests/csharp/test_csharp_index_from_end.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = { 10, 20, 30 };
__Check((data[^1]).ToString(), "30");
