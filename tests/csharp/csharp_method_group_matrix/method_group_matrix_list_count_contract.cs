// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
var values = new System.Collections.Generic.List<int> { 79, 80, 79 }; __Check((values.Count == 3).ToString(), "True");
