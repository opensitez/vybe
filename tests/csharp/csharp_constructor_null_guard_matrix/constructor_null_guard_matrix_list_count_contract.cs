// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
var values = new System.Collections.Generic.List<int> { 126, 127, 126 }; __Check((values.Count == 3).ToString(), "True");
