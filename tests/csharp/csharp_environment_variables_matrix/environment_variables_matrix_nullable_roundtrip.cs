// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
int? maybe = 100; __Check((maybe.HasValue && maybe.Value == 100).ToString(), "True");
