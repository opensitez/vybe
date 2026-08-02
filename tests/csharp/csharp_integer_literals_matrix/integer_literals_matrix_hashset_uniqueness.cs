// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(15); set.Add(15); __Check((set.Count == 1).ToString(), "True");
