// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
int seed = 92; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
