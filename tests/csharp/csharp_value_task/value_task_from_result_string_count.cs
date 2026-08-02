// vybe-test: csharp/csharp_value_task/value_task_from_result_string_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Run() {
    var vt = System.Threading.Tasks.ValueTask.FromResult("abc");
    string s = await vt;
    __Check((s.Length).ToString(), "3");
}
Run().Wait();
