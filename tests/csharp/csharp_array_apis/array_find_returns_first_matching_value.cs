// vybe-test: csharp/csharp_array_apis/array_find_returns_first_matching_value
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var values = new[] { 2, 4, 5, 8 }; __Check((System.Array.Find(values, value => value % 2 == 1)).ToString(), "5");
