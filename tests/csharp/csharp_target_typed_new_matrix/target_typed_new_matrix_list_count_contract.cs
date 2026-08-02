// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// target_typed_new_matrix
var values = new System.Collections.Generic.List<int> { 107, 108, 107 }; __Check((values.Count == 3).ToString(), "True");
