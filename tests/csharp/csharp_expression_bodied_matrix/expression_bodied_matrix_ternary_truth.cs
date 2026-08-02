// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
int seed = 106; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
