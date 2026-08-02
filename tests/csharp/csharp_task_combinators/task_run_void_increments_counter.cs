// vybe-test: csharp/csharp_task_combinators/task_run_void_increments_counter
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    int count = 0;
    System.Threading.Tasks.Task.Run(() => { count = 4; }).Wait();
    __Check((count).ToString(), "4");
}
Run().Wait();
