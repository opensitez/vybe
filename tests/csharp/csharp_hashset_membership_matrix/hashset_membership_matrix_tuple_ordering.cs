// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
var tuple = (left: 33, right: 34); __Check((tuple.left < tuple.right).ToString(), "True");
