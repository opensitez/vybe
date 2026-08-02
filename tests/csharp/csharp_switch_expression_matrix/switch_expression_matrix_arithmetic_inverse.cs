// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
int seed = 43; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
