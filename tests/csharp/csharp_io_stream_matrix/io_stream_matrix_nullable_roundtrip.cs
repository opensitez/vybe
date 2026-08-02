// vybe-test: csharp/csharp_io_stream_matrix/io_stream_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_io_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_stream_matrix
int? maybe = 90; __Check((maybe.HasValue && maybe.Value == 90).ToString(), "True");
