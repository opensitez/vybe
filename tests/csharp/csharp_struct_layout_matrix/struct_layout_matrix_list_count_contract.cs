// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
var values = new System.Collections.Generic.List<int> { 113, 114, 113 }; __Check((values.Count == 3).ToString(), "True");
