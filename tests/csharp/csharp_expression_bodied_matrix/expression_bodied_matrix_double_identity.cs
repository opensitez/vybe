// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
double seed = 106; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
