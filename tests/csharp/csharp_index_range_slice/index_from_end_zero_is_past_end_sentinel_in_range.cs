// vybe-test: csharp/csharp_index_range_slice/index_from_end_zero_is_past_end_sentinel_in_range
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3}; var slice=data[1..^0]; __Check((slice.Length).ToString(), "2");
