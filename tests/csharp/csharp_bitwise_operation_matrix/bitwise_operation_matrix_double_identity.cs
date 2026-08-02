// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// bitwise_operation_matrix
double seed = 104; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
