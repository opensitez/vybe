// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_coalescing_matrix
var values = new System.Collections.Generic.List<int> { 56, 57, 56 }; __Check((values.Count == 3).ToString(), "True");
