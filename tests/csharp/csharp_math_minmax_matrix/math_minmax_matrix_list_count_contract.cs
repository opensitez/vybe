// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_minmax_matrix
var values = new System.Collections.Generic.List<int> { 101, 102, 101 }; __Check((values.Count == 3).ToString(), "True");
