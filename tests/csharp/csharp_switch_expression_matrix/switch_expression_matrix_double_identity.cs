// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
double seed = 43; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
