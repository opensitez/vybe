// vybe-test: csharp/csharp_index_range_slice/array_slice_all_but_last
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={9,8,7,6}; var slice=data[..^1]; __Check((string.Join(",",slice)).ToString(), "9,8,7");
