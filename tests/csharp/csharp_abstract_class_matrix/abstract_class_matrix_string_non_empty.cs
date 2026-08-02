// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
string feature = "abstract_class_matrix"; __Check((feature.Length > 0).ToString(), "True");
