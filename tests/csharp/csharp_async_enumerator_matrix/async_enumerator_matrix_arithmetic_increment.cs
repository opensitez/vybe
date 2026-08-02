// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_enumerator_matrix
int seed = 116; __Check((seed + 1 > seed).ToString(), "True");
