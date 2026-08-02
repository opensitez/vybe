// vybe-test: csharp/csharp_task_completion_matrix/task_completion_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_task_completion_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_completion_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(85); set.Add(85); __Check((set.Count == 1).ToString(), "True");
