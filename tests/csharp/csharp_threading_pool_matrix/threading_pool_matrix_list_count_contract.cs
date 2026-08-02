// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
var values = new System.Collections.Generic.List<int> { 87, 88, 87 }; __Check((values.Count == 3).ToString(), "True");
