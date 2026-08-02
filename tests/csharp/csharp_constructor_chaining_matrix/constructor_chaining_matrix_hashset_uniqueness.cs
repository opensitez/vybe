// vybe-test: csharp/csharp_constructor_chaining_matrix/constructor_chaining_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chaining_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_chaining_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(68); set.Add(68); __Check((set.Count == 1).ToString(), "True");
