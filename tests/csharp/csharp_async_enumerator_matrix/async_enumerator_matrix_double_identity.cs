// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_enumerator_matrix
double seed = 116; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
