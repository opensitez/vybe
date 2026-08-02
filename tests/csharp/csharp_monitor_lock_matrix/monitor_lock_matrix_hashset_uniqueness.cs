// vybe-test: csharp/csharp_monitor_lock_matrix/monitor_lock_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_monitor_lock_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// monitor_lock_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(84); set.Add(84); __Check((set.Count == 1).ToString(), "True");
