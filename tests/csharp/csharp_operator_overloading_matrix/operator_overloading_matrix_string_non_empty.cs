// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
string feature = "operator_overloading_matrix"; __Check((feature.Length > 0).ToString(), "True");
