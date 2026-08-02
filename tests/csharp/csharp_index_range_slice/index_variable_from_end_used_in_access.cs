// vybe-test: csharp/csharp_index_range_slice/index_variable_from_end_used_in_access
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={5,10,15}; System.Index idx=^2; __Check((data[idx]).ToString(), "10");
