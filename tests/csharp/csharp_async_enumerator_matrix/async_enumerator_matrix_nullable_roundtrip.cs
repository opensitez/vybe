// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_enumerator_matrix
int? maybe = 116; __Check((maybe.HasValue && maybe.Value == 116).ToString(), "True");
