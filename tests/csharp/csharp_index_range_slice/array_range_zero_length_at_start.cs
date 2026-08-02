// vybe-test: csharp/csharp_index_range_slice/array_range_zero_length_at_start
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3}; var slice=data[0..0]; __Check((slice.Length).ToString(), "0");
