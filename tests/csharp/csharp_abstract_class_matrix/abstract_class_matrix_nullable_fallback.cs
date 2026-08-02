// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
int? maybe = null; int fallback = maybe ?? 72; __Check((fallback == 72).ToString(), "True");
