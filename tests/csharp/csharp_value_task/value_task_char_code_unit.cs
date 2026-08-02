// vybe-test: csharp/csharp_value_task/value_task_char_code_unit
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<char> Get() { return 'Z'; }
async System.Threading.Tasks.Task Run() {
    char c = await Get();
    __Check(((int)c).ToString(), "90");
}
Run().Wait();
