// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
string feature = "boxing_unboxing_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
