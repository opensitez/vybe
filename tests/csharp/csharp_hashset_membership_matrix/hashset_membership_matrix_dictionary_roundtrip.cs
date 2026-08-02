// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[33] = 34; __Check((map.ContainsKey(33) && map[33] == 34).ToString(), "True");
