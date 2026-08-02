// vybe-test: csharp/csharp_index_range_slice/range_full_array_via_explicit_bounds
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={4,5,6}; var slice=data[0..3]; __Check((slice.Length).ToString(), "3"); __Check((slice[2]).ToString(), "6");
