// vybe-test: csharp/csharp_task_combinators/when_any_continue_with_doubles_winner
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Win() { return 6; }
async System.Threading.Tasks.Task<int> Lose() {
    await System.Threading.Tasks.Task.Delay(300);
    return 1;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    var winner = await System.Threading.Tasks.Task.WhenAny(Win(), Lose());
    await winner.ContinueWith(t => count = t.Result * 2);
    __Check((count).ToString(), "12");
}
Run().Wait();
