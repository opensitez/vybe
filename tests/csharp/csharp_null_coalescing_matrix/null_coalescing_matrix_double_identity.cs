// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_coalescing_matrix
double seed = 56; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
