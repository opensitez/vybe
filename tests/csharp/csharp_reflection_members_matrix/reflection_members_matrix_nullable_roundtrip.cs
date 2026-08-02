// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
int? maybe = 92; __Check((maybe.HasValue && maybe.Value == 92).ToString(), "True");
