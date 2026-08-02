// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
var values = new System.Collections.Generic.List<int> { 119, 120, 119 }; __Check((values.Count == 3).ToString(), "True");
