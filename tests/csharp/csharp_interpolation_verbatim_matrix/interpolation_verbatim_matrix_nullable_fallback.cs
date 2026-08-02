// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
int? maybe = null; int fallback = maybe ?? 110; __Check((fallback == 110).ToString(), "True");
