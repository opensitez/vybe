// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
int seed = 106; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
