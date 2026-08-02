// vybe-test: csharp/csharp_async_state_machine_matrix/async_state_machine_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_async_state_machine_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_state_machine_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(88); set.Add(88); __Check((set.Count == 1).ToString(), "True");
