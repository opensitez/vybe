// vybe-test: csharp/csharp_value_task/value_task_empty_string_length_zero
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<string> Empty() { return ""; }
async System.Threading.Tasks.Task Run() {
    string s = await Empty();
    __Check((s.Length).ToString(), "0");
}
Run().Wait();
