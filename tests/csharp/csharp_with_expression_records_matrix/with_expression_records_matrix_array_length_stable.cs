// vybe-test: csharp/csharp_with_expression_records_matrix/with_expression_records_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_records_matrix
int seed = 109; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
