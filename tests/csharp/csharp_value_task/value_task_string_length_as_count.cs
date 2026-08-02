// vybe-test: csharp/csharp_value_task/value_task_string_length_as_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<string> Name() { return "hello"; }
async System.Threading.Tasks.Task Run() {
    string s = await Name();
    __Check((s.Length).ToString(), "5");
}
Run().Wait();
