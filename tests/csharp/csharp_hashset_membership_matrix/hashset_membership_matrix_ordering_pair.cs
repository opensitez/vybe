// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
int seed = 33; int right = seed + 1; __Check((seed < right).ToString(), "True");
