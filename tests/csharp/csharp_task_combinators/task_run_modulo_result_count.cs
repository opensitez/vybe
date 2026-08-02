// vybe-test: csharp/csharp_task_combinators/task_run_modulo_result_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => 17 % 5);
    __Check((t.Result).ToString(), "2");
}
Run().Wait();
