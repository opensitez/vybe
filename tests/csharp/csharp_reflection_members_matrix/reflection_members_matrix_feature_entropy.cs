// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
string feature = "reflection_members_matrix:92"; __Check((feature.Length >= 1).ToString(), "True");
