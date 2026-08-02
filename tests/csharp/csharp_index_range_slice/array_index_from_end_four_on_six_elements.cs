// vybe-test: csharp/csharp_index_range_slice/array_index_from_end_four_on_six_elements
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={10,20,30,40,50,60}; __Check((data[^4]).ToString(), "30");
