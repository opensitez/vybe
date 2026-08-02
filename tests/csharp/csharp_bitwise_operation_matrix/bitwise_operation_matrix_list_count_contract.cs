// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// bitwise_operation_matrix
var values = new System.Collections.Generic.List<int> { 104, 105, 104 }; __Check((values.Count == 3).ToString(), "True");
