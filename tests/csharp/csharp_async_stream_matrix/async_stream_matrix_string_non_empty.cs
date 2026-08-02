// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_stream_matrix
string feature = "async_stream_matrix"; __Check((feature.Length > 0).ToString(), "True");
