// vybe-test: csharp/csharp_task_combinators/continue_with_chained_two_steps
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(2)
        .ContinueWith(t => t.Result + 3)
        .ContinueWith(t => count = t.Result * 4);
    __Check((count).ToString(), "20");
}
Run().Wait();
