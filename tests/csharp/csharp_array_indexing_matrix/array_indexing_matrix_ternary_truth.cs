// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
int seed = 24; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
