// vybe-test: csharp/csharp_index_range_slice/range_half_open_slice_length_computed
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3,4,5,6,7}; var slice=data[2..5]; __Check((slice.Length).ToString(), "3");
