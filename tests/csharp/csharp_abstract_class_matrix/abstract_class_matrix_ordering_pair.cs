// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
int seed = 72; int right = seed + 1; __Check((seed < right).ToString(), "True");
