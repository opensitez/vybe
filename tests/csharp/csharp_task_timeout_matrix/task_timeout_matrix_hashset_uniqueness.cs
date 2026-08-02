// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(86); set.Add(86); __Check((set.Count == 1).ToString(), "True");
