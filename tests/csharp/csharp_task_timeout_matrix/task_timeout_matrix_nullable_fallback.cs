// vybe-test: csharp/csharp_task_timeout_matrix/task_timeout_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_task_timeout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// task_timeout_matrix
int? maybe = null; int fallback = maybe ?? 86; __Check((fallback == 86).ToString(), "True");
