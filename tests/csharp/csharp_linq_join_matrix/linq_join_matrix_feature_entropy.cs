// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
string feature = "linq_join_matrix:119"; __Check((feature.Length >= 1).ToString(), "True");
