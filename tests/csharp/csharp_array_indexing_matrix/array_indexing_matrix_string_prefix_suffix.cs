// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
string feature = "array_indexing_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
