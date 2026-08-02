// vybe-test: csharp/csharp_index_range_slice/array_slice_second_half_length
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3,4}; var slice=data[2..4]; __Check((slice[0]).ToString(), "3"); __Check((slice[1]).ToString(), "4");
