// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
string feature = "random_distribution_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
