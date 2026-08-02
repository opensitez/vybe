// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
int seed = 92; int right = seed + 1; __Check((seed < right).ToString(), "True");
