// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
int seed = 119; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
