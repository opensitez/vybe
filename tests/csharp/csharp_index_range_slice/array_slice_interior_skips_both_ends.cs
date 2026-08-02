// vybe-test: csharp/csharp_index_range_slice/array_slice_interior_skips_both_ends
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={0,1,2,3,4}; var slice=data[1..4]; __Check((slice[0]).ToString(), "1"); __Check((slice[2]).ToString(), "3");
