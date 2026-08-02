// vybe-test: csharp/csharp_task_combinators/continue_with_doubles_task_result
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.Run(() => 7)
        .ContinueWith(t => count = t.Result * 2);
    __Check((count).ToString(), "14");
}
Run().Wait();
