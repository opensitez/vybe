// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
int seed = 98; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
