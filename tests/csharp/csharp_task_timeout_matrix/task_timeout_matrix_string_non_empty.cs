// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
string feature = "task_timeout_matrix"; __Check((feature.Length > 0).ToString(), "True");
