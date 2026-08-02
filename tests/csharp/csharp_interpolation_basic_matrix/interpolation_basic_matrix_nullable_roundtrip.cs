// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
int? maybe = 112; __Check((maybe.HasValue && maybe.Value == 112).ToString(), "True");
