// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
int seed = 54; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
