// vybe-test: csharp/csharp_task_combinators/task_run_string_length_as_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => "hello".Length);
    __Check((t.Result).ToString(), "5");
}
Run().Wait();
