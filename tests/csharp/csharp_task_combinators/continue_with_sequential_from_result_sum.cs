// vybe-test: csharp/csharp_task_combinators/continue_with_sequential_from_result_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(3)
        .ContinueWith(t => count += t.Result);
    await System.Threading.Tasks.Task.FromResult(4)
        .ContinueWith(t => count += t.Result);
    __Check((count).ToString(), "7");
}
Run().Wait();
