// vybe-test: csharp/csharp_do_while_matrix/do_while_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_do_while_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// do_while_matrix
double seed = 48; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
