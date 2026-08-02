// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_stream_matrix
int seed = 111; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
