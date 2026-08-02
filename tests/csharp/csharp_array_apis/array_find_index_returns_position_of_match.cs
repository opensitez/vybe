// vybe-test: csharp/csharp_array_apis/array_find_index_returns_position_of_match
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var values = new[] { 2, 4, 5, 8 }; __Check((System.Array.FindIndex(values, value => value % 2 == 1)).ToString(), "2");
