// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// bitwise_operation_matrix
int? maybe = null; int fallback = maybe ?? 104; __Check((fallback == 104).ToString(), "True");
