// vybe-test: csharp/csharp_index_range_slice/range_single_element_slice_length_one
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={9,8,7}; var slice=data[1..2]; __Check((slice.Length).ToString(), "1"); __Check((slice[0]).ToString(), "8");
