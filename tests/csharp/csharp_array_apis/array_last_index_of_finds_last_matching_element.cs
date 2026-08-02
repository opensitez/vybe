// vybe-test: csharp/csharp_array_apis/array_last_index_of_finds_last_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var values = new[] { 1, 2, 1, 3 }; __Check((System.Array.LastIndexOf(values, 1)).ToString(), "2");
