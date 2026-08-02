// vybe-test: csharp/csharp_index_range_slice/index_from_end_on_empty_range_start_marker
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2}; var slice=data[2..]; __Check((slice.Length).ToString(), "0");
