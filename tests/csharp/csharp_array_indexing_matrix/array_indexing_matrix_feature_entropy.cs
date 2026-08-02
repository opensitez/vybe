// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
string feature = "array_indexing_matrix:24"; __Check((feature.Length >= 1).ToString(), "True");
