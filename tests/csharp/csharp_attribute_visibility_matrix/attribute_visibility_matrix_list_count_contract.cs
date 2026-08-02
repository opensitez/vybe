// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
var values = new System.Collections.Generic.List<int> { 93, 94, 93 }; __Check((values.Count == 3).ToString(), "True");
