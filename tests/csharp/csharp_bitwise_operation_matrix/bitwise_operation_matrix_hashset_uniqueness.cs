// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// bitwise_operation_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(104); set.Add(104); __Check((set.Count == 1).ToString(), "True");
