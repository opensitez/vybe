// vybe-test: csharp/csharp_task_combinators/continue_with_from_when_any_result
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> A() { return 4; }
async System.Threading.Tasks.Task<int> B() { return 8; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    var winner = await System.Threading.Tasks.Task.WhenAny(A(), B());
    await winner.ContinueWith(t => count = t.Result + 1);
    __Check((count).ToString(), "5");
}
Run().Wait();
