// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
double seed = 119; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
