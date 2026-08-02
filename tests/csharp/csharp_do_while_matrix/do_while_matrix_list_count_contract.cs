// vybe-test: csharp/csharp_do_while_matrix/do_while_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_do_while_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// do_while_matrix
var values = new System.Collections.Generic.List<int> { 48, 49, 48 }; __Check((values.Count == 3).ToString(), "True");
