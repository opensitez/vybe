// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_coalescing_matrix
int? maybe = 56; __Check((maybe.HasValue && maybe.Value == 56).ToString(), "True");
