// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
string feature = "hashset_membership_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
