// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
var tuple = (left: 86, right: 87); __Check((tuple.left < tuple.right).ToString(), "True");
