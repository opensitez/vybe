// vybe-test: csharp/csharp_do_while_matrix/do_while_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_do_while_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// do_while_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(48); set.Add(48); __Check((set.Count == 1).ToString(), "True");
