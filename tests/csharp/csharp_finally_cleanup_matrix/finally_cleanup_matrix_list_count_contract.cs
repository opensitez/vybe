// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
var values = new System.Collections.Generic.List<int> { 54, 55, 54 }; __Check((values.Count == 3).ToString(), "True");
