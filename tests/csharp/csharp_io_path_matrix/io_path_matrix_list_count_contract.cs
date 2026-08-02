// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
var values = new System.Collections.Generic.List<int> { 122, 123, 122 }; __Check((values.Count == 3).ToString(), "True");
