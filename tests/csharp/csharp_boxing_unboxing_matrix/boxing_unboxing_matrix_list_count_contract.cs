// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
var values = new System.Collections.Generic.List<int> { 62, 63, 62 }; __Check((values.Count == 3).ToString(), "True");
