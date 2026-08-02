// vybe-test: csharp/csharp_task_combinators/when_any_loser_not_awaited_count_stays_one
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Win() { return 1; }
async System.Threading.Tasks.Task<int> Lose() {
    await System.Threading.Tasks.Task.Delay(500);
    return 99;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Win(), Lose());
    int count = winner.Result;
    __Check((count).ToString(), "1");
}
Run().Wait();
