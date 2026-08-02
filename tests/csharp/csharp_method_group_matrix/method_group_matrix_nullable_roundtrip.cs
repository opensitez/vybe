// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
int? maybe = 79; __Check((maybe.HasValue && maybe.Value == 79).ToString(), "True");
