// vybe-test: csharp/csharp_pointer_like_emulation_matrix/pointer_like_emulation_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_pointer_like_emulation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pointer_like_emulation_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(114); set.Add(114); __Check((set.Count == 1).ToString(), "True");
