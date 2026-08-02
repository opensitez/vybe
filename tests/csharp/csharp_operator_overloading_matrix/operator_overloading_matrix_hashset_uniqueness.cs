// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(105); set.Add(105); __Check((set.Count == 1).ToString(), "True");
