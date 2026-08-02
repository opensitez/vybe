// vybe-test: csharp/csharp_index_range_slice/range_on_long_array_large_offset
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={0,1,2,3,4,5,6,7,8,9}; var slice=data[7..]; __Check((slice[0]).ToString(), "7"); __Check((slice[2]).ToString(), "9");
