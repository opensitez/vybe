// vybe-test: csharp/csharp_array_apis/array_index_of_finds_first_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var values = new[] { 4, 7, 9 }; __Check((System.Array.IndexOf(values, 7)).ToString(), "1");
