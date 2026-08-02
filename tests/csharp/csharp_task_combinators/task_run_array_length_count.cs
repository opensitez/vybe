// vybe-test: csharp/csharp_task_combinators/task_run_array_length_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => new int[] { 1, 2, 3, 4 }.Length);
    __Check((t.Result).ToString(), "4");
}
Run().Wait();
