// vybe-test: csharp/csharp_index_range_slice/array_slice_assign_is_independent_copy
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3}; var slice=data[0..2]; slice[0]=9; __Check((data[0]).ToString(), "1"); __Check((slice[0]).ToString(), "9");
