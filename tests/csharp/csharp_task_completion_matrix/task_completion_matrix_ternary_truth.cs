// vybe-test: csharp/csharp_task_completion_matrix/task_completion_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_task_completion_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_completion_matrix
int seed = 85; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
