// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_constraints_matrix
int? maybe = 80; __Check((maybe.HasValue && maybe.Value == 80).ToString(), "True");
