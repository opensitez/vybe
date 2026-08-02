// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_stream_matrix
int? maybe = 111; __Check((maybe.HasValue && maybe.Value == 111).ToString(), "True");
