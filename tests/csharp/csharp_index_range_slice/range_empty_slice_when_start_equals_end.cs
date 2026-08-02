// vybe-test: csharp/csharp_index_range_slice/range_empty_slice_when_start_equals_end
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3}; var slice=data[2..2]; __Check((slice.Length).ToString(), "0");
