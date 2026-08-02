// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
int seed = 92; __Check((seed - seed == 0).ToString(), "True");
