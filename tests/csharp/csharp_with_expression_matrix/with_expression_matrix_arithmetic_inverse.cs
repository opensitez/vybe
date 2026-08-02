// vybe-test: csharp/csharp_with_expression_matrix/with_expression_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_matrix
int seed = 108; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
