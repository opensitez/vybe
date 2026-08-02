// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
double seed = 72; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
