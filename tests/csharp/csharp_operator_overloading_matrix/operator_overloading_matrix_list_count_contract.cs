// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
var values = new System.Collections.Generic.List<int> { 105, 106, 105 }; __Check((values.Count == 3).ToString(), "True");
