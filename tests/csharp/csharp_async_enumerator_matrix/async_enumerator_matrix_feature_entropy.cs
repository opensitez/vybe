// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_enumerator_matrix
string feature = "async_enumerator_matrix:116"; __Check((feature.Length >= 1).ToString(), "True");
