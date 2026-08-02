// vybe-test: csharp/csharp_async_task/awaiting_task_delay_resumes_execution_after_pause
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    __Check(("before").ToString(), "before");
    await System.Threading.Tasks.Task.Delay(1);
    __Check(("after").ToString(), "after");
}
Run().Wait();
