// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
double seed = 33; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
