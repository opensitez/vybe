// vybe-test: csharp/csharp_async_task_matrix/task_completed_task_reports_completion
// origin: languages/csharp/tests/csharp/test_csharp_async_task_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var completed = System.Threading.Tasks.Task.CompletedTask;
__Check((completed.IsCompleted).ToString(), "True");
__Check((completed.IsFaulted).ToString(), "False");
