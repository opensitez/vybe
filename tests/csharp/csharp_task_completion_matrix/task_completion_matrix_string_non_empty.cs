// vybe-test: csharp/csharp_task_completion_matrix/task_completion_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_task_completion_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_completion_matrix
string feature = "task_completion_matrix"; __Check((feature.Length > 0).ToString(), "True");
