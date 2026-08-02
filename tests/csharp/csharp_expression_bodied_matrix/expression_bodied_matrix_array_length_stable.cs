// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
int seed = 106; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
