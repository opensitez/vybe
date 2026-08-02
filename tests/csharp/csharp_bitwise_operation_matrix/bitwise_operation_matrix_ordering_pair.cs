// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// bitwise_operation_matrix
int seed = 104; int right = seed + 1; __Check((seed < right).ToString(), "True");
