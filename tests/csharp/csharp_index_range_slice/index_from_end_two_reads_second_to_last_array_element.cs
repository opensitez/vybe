// vybe-test: csharp/csharp_index_range_slice/index_from_end_two_reads_second_to_last_array_element
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={10,20,30,40}; __Check((data[^2]).ToString(), "30");
