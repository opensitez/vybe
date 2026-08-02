// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// delegate_invocation_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(74); set.Add(74); __Check((set.Count == 1).ToString(), "True");
