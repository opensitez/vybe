// vybe-test: csharp/csharp_task_completion_matrix/task_completion_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_task_completion_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_completion_matrix
double seed = 85; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
