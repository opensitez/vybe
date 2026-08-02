// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
int? maybe = null; int fallback = maybe ?? 92; __Check((fallback == 92).ToString(), "True");
