// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
var values = new System.Collections.Generic.List<int> { 72, 73, 72 }; __Check((values.Count == 3).ToString(), "True");
