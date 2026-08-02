// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
int? maybe = 110; __Check((maybe.HasValue && maybe.Value == 110).ToString(), "True");
