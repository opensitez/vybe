// vybe-test: csharp/csharp_async_task/exception_in_async_method_propagates_through_await
// origin: languages/csharp/tests/csharp/test_csharp_async_task.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

async System.Threading.Tasks.Task Fail() {
    await System.Threading.Tasks.Task.Yield();
    throw new System.Exception("async fail");
}
string msg = "";
try { Fail().Wait(); }
catch (System.AggregateException ae) { msg = ae.InnerException.Message; }
__Check((msg).ToString(), "async fail");
