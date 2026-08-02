// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
string feature = "expression_bodied_matrix"; __Check((feature.Length > 0).ToString(), "True");
