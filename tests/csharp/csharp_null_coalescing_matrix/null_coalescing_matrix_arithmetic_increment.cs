// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_coalescing_matrix
int seed = 56; __Check((seed + 1 > seed).ToString(), "True");
