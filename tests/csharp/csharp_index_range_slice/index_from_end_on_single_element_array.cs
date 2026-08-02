// vybe-test: csharp/csharp_index_range_slice/index_from_end_on_single_element_array
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={99}; __Check((data[^1]).ToString(), "99");
