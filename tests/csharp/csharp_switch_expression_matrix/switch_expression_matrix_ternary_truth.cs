// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
int seed = 43; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
