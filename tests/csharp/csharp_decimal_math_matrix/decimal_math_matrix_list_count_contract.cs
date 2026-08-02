// vybe-test: csharp/csharp_decimal_math_matrix/decimal_math_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_decimal_math_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// decimal_math_matrix
var values = new System.Collections.Generic.List<int> { 17, 18, 17 }; __Check((values.Count == 3).ToString(), "True");
