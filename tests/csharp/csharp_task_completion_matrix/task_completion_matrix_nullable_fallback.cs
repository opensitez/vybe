// vybe-test: csharp/csharp_task_completion_matrix/task_completion_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_task_completion_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_completion_matrix
int? maybe = null; int fallback = maybe ?? 85; __Check((fallback == 85).ToString(), "True");
