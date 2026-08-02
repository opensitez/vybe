// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
int seed = 43; int right = seed + 1; __Check((seed < right).ToString(), "True");
