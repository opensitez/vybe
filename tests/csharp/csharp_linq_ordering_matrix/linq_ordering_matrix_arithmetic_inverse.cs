// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_ordering_matrix
int seed = 121; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
