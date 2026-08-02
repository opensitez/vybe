// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
int? maybe = null; int fallback = maybe ?? 112; __Check((fallback == 112).ToString(), "True");
