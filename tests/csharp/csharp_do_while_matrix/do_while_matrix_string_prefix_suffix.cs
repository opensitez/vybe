// vybe-test: csharp/csharp_do_while_matrix/do_while_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_do_while_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// do_while_matrix
string feature = "do_while_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
