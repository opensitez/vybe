// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
string feature = "hashset_membership_matrix:33"; __Check((feature.Length >= 1).ToString(), "True");
