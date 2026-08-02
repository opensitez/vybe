// vybe-test: csharp/csharp_value_task/value_task_exception_message_length_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Fail() {
    throw new System.Exception("err");
}
async System.Threading.Tasks.Task Run() {
    int len = 0;
    try { await Fail(); }
    catch (System.Exception ex) { len = ex.Message.Length; }
    __Check((len).ToString(), "3");
}
Run().Wait();
