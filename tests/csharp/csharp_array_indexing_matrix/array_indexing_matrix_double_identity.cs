// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
double seed = 24; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
