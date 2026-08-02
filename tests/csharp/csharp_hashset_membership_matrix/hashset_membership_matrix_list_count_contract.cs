// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
var values = new System.Collections.Generic.List<int> { 33, 34, 33 }; __Check((values.Count == 3).ToString(), "True");
