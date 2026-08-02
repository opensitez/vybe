// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
string feature = "reflection_members_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
