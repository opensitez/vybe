// vybe-test: csharp/csharp_task_combinators/when_all_from_result_tasks_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        System.Threading.Tasks.Task.FromResult(7),
        System.Threading.Tasks.Task.FromResult(8)
    );
    __Check((results[0] + results[1]).ToString(), "15");
}
Run().Wait();
