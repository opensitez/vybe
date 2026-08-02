// vybe-test: csharp/csharp_array_apis/array_binary_search_returns_position_of_found_value
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var values = new[] { 1, 3, 5, 7 }; __Check((System.Array.BinarySearch(values, 5)).ToString(), "2");
