// vybe-test: csharp/csharp_array_binary_search/array_binary_search_returns_complement_for_missing_insertion_point
// origin: languages/csharp/tests/csharp/test_csharp_array_binary_search.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] sorted = { 2, 4, 8 };
int index = System.Array.BinarySearch(sorted, 5);
__Check((index < 0).ToString(), "True");
__Check((~index).ToString(), "2");
