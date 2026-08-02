// vybe-test: csharp/csharp_task_combinators/task_run_bool_true_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => true);
    __Check((t.Result ? 1 : 0).ToString(), "1");
}
Run().Wait();
