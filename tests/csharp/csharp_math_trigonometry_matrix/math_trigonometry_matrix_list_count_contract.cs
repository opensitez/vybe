// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
var values = new System.Collections.Generic.List<int> { 102, 103, 102 }; __Check((values.Count == 3).ToString(), "True");
