// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
int seed = 55; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
