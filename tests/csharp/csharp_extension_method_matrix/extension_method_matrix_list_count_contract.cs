// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
var values = new System.Collections.Generic.List<int> { 78, 79, 78 }; __Check((values.Count == 3).ToString(), "True");
