// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
int seed = 106; int right = seed + 1; __Check((seed < right).ToString(), "True");
