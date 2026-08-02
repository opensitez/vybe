// vybe-test: csharp/csharp_value_task/value_task_null_string_length_zero
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<string> NullStr() { return null; }
async System.Threading.Tasks.Task Run() {
    string s = await NullStr();
    __Check((s == null ? 0 : s.Length).ToString(), "0");
}
Run().Wait();
