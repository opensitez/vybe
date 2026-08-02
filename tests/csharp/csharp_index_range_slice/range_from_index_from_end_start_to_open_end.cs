// vybe-test: csharp/csharp_index_range_slice/range_from_index_from_end_start_to_open_end
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3,4,5}; var slice=data[^3..]; __Check((slice.Length).ToString(), "3"); __Check((slice[0]).ToString(), "3");
