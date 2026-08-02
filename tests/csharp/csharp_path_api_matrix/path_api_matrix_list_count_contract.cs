// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
var values = new System.Collections.Generic.List<int> { 123, 124, 123 }; __Check((values.Count == 3).ToString(), "True");
