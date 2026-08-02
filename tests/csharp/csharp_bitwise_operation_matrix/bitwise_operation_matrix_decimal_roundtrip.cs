// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// bitwise_operation_matrix
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
