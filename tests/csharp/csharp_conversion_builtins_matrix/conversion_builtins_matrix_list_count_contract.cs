// vybe-test: csharp/csharp_conversion_builtins_matrix/conversion_builtins_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_conversion_builtins_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// conversion_builtins_matrix
var values = new System.Collections.Generic.List<int> { 124, 125, 124 }; __Check((values.Count == 3).ToString(), "True");
