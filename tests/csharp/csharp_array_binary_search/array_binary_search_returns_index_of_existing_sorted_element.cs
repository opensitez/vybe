// vybe-test: csharp/csharp_array_binary_search/array_binary_search_returns_index_of_existing_sorted_element
// origin: languages/csharp/tests/csharp/test_csharp_array_binary_search.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] sorted = { 2, 4, 6, 8 };
__Check((System.Array.BinarySearch(sorted, 6)).ToString(), "2");
