// vybe-test: csharp/csharp_index_range_slice/index_from_end_three_reads_third_from_end
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3,4,5,6}; __Check((data[^3]).ToString(), "4");
