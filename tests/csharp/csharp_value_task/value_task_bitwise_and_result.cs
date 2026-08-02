// vybe-test: csharp/csharp_value_task/value_task_bitwise_and_result
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.ValueTask<int> Mask() { return 0xF0; }
async System.Threading.Tasks.ValueTask<int> Value() { return 0xFF; }
async System.Threading.Tasks.Task Run() {
    __Check((await Mask() & await Value()).ToString(), "240");
}
Run().Wait();
