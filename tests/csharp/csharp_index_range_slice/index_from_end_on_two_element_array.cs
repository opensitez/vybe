// vybe-test: csharp/csharp_index_range_slice/index_from_end_on_two_element_array
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={7,8}; __Check((data[^1]).ToString(), "8"); __Check((data[^2]).ToString(), "7");
