// vybe-test: csharp/csharp_ranges_indices/index_from_end_one_is_last_element
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a={1,2,3,4,5}; __Check((a[^1]).ToString(), "5");
