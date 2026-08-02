// vybe-test: csharp/csharp_task_combinators/when_any_two_identical_fast_tasks_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Fast() { return 4; }
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Fast(), Fast());
    __Check((winner.Result).ToString(), "4");
}
Run().Wait();
