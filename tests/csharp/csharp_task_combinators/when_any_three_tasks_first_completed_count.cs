// vybe-test: csharp/csharp_task_combinators/when_any_three_tasks_first_completed_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> A() { return 10; }
async System.Threading.Tasks.Task<int> B() { return 20; }
async System.Threading.Tasks.Task<int> C() { return 30; }
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(A(), B(), C());
    __Check((winner.Result).ToString(), "10");
}
Run().Wait();
