// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_stream_matrix
int seed = 111; __Check((seed + 1 > seed).ToString(), "True");
