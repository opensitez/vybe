// vybe-test: csharp/csharp_task_advanced/task_continuation_runs_after_completion
// origin: languages/csharp/tests/csharp/test_csharp_task_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int result=0;
System.Threading.Tasks.Task.Run(()=>7)
    .ContinueWith(t=>result=t.Result*2)
    .Wait();
__Check((result).ToString(), "14");
