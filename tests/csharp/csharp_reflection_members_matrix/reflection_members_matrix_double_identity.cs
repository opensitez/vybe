// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
double seed = 92; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
