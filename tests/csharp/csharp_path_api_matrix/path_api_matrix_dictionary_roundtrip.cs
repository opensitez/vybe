// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[123] = 124; __Check((map.ContainsKey(123) && map[123] == 124).ToString(), "True");
