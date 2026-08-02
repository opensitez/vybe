// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
string feature = "method_group_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
