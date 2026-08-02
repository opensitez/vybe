// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_constraints_matrix
var values = new System.Collections.Generic.List<int> { 80, 81, 80 }; __Check((values.Count == 3).ToString(), "True");
