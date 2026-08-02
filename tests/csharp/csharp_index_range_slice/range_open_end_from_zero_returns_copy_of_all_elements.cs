// vybe-test: csharp/csharp_index_range_slice/range_open_end_from_zero_returns_copy_of_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={11,22,33}; var slice=data[0..]; __Check((slice.Length).ToString(), "3");
