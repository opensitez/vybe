// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_stream_matrix
string feature = "async_stream_matrix:111"; __Check((feature.Length >= 1).ToString(), "True");
