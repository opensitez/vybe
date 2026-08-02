// vybe-test: csharp/csharp_task_combinators/task_run_three_parallel_workers_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var a = System.Threading.Tasks.Task.Run(() => 1);
    var b = System.Threading.Tasks.Task.Run(() => 2);
    var c = System.Threading.Tasks.Task.Run(() => 3);
    __Check((a.Result + b.Result + c.Result).ToString(), "6");
}
Run().Wait();
