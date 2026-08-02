// vybe-test: csharp/csharp_index_range_slice/array_slice_first_element_preserved
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={100,200,300}; var slice=data[..1]; __Check((slice[0]).ToString(), "100");
