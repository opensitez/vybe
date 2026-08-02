// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
var values = new System.Collections.Generic.List<int> { 103, 104, 103 }; __Check((values.Count == 3).ToString(), "True");
