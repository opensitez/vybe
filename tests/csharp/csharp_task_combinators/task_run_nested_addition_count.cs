// vybe-test: csharp/csharp_task_combinators/task_run_nested_addition_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var outer = System.Threading.Tasks.Task.Run(() => {
        return System.Threading.Tasks.Task.Run(() => 2 + 3).Result;
    });
    __Check((outer.Result).ToString(), "5");
}
Run().Wait();
