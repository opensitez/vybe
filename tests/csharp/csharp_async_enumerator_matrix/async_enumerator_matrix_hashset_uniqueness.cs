// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_enumerator_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(116); set.Add(116); __Check((set.Count == 1).ToString(), "True");
