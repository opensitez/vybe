// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
int seed = 98; __Check((seed + 1 > seed).ToString(), "True");
