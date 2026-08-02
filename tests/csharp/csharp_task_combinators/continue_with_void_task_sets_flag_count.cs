// vybe-test: csharp/csharp_task_combinators/continue_with_void_task_sets_flag_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.Run(() => { })
        .ContinueWith(t => count = 1);
    __Check((count).ToString(), "1");
}
Run().Wait();
