// vybe-test: csharp/csharp_async_task/async_method_returns_task_t_with_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task<int> Compute() {
    await System.Threading.Tasks.Task.Yield();
    return 7;
}
__Check((Compute().Result).ToString(), "7");
