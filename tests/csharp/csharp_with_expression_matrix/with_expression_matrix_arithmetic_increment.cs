// vybe-test: csharp/csharp_with_expression_matrix/with_expression_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_matrix
int seed = 108; __Check((seed + 1 > seed).ToString(), "True");
