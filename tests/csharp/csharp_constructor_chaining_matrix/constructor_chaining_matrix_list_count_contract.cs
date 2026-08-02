// vybe-test: csharp/csharp_constructor_chaining_matrix/constructor_chaining_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chaining_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_chaining_matrix
var values = new System.Collections.Generic.List<int> { 68, 69, 68 }; __Check((values.Count == 3).ToString(), "True");
