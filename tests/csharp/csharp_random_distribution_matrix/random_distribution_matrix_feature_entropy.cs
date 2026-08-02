// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
string feature = "random_distribution_matrix:98"; __Check((feature.Length >= 1).ToString(), "True");
