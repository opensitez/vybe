// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
double seed = 54; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
