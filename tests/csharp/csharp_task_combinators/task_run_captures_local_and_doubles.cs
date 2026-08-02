// vybe-test: csharp/csharp_task_combinators/task_run_captures_local_and_doubles
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    int seed = 5;
    var t = System.Threading.Tasks.Task.Run(() => seed * 2);
    __Check((t.Result).ToString(), "10");
}
Run().Wait();
