// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
int? maybe = null; int fallback = maybe ?? 105; __Check((fallback == 105).ToString(), "True");
