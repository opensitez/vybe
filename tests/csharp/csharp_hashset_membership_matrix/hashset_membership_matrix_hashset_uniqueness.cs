// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(33); set.Add(33); __Check((set.Count == 1).ToString(), "True");
