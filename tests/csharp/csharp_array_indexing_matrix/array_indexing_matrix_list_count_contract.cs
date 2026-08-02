// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
var values = new System.Collections.Generic.List<int> { 24, 25, 24 }; __Check((values.Count == 3).ToString(), "True");
