// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_enumerator_matrix
int seed = 116; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
