// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
var values = new System.Collections.Generic.List<int> { 50, 51, 50 }; __Check((values.Count == 3).ToString(), "True");
