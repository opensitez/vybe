// vybe-test: csharp/csharp_with_expression_records_matrix/with_expression_records_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_records_matrix
int seed = 109; int right = seed + 1; __Check((seed < right).ToString(), "True");
