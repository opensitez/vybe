// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
string feature = "switch_expression_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
