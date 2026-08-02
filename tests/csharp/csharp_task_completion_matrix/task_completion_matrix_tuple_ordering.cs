// vybe-test: csharp/csharp_task_completion_matrix/task_completion_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_task_completion_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_completion_matrix
var tuple = (left: 85, right: 86); __Check((tuple.left < tuple.right).ToString(), "True");
