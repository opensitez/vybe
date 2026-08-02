// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(92); set.Add(92); __Check((set.Count == 1).ToString(), "True");
