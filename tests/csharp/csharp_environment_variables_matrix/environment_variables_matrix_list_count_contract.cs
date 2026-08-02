// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
var values = new System.Collections.Generic.List<int> { 100, 101, 100 }; __Check((values.Count == 3).ToString(), "True");
