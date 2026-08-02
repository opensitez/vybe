// vybe-test: csharp/csharp_linq_groupby_matrix/linq_groupby_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_groupby_matrix
string feature = "linq_groupby_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
