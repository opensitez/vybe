// vybe-test: csharp/csharp_task_advanced/task_run_executes_lambda_on_thread_pool
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t=System.Threading.Tasks.Task.Run(()=>42);
__Check((t.Result).ToString(), "42");
