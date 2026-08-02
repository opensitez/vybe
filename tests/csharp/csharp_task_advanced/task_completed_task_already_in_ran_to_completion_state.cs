// vybe-test: csharp/csharp_task_advanced/task_completed_task_already_in_ran_to_completion_state
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t=System.Threading.Tasks.Task.CompletedTask;
__Check((t.IsCompleted).ToString(), "True");
